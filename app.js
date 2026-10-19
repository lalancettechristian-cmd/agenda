/* Agenda — app iPhone du Cahier à spirale.

   Copie locale de l'agenda (IndexedDB), utilisable hors ligne, synchronisée avec cahier.json dans OneDrive par
   Microsoft Graph (connexion Microsoft : OAuth 2.0 avec PKCE). Même format que l'application Windows (données en
   version 8) : l'iPhone est un poste de plus (« iPhone » dans modifie_par). Les champs inconnus sont gardés tels
   quels : l'app modifie l'objet lu, elle ne le reconstruit jamais.

   Envoi : PUT avec If-Match (l'eTag de la version reçue). Si le PC a modifié l'agenda entre-temps, les deux versions
   sont fusionnées quand elles touchent des pages différentes ; sinon, la version de l'iPhone est gardée à part
   (cahier-conflit-iPhone-date.json, comme le fait le PC) et l'iPhone prend celle du PC. Le PC relit seul cahier.json
   quand il change (toutes les 4 s) : l'iPhone peut enregistrer même quand l'agenda est ouvert sur le PC. */
(function () {
"use strict";

const CONFIG = Object.assign({clientId: "", autorite: "consumers", dossier: "Agenda", redirection: "", base: "agenda-iphone"},
  window.AGENDA_CONFIG || {});
const VERSION_APP = "1.0.3";
const VERSION_DONNEES = 8;
const DUREE_VERROU = 180;          // secondes sans battement : le verrou d'un poste est abandonné
const OBJ = "￼";              // une image dans le texte d'une page
const LIGNES_JOUR = 3;             // lignes par jour dans l'agenda local
const NB_CASES = 29;               // lignes de cases d'une page en colonnes (comme sur le PC)
const FICHIER = "cahier.json";
const POSTE = "iPhone";
const PORTEES = "Files.ReadWrite User.Read offline_access";
const GRAPH = "https://graph.microsoft.com/v1.0";
const CONNEXION = "https://login.microsoftonline.com/";
const CACHE_IMAGES = "images-agenda";

const MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre",
  "novembre", "décembre"];
const MOIS_ABR = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."];
const JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"];
const JOURS_COURTS = ["L", "M", "M", "J", "V", "S", "D"];
const SERVICES = {
  agenda: ["Google Agenda", "https://calendar.google.com/"], todo: ["Google Tasks", "https://tasks.google.com/"],
  photo: ["Google Photos", "https://photos.google.com/"], courriels: ["Gmail", "https://mail.google.com/"],
  contacts: ["Google Contacts", "https://contacts.google.com/"], page: ["Page web", "https://www.google.com/"],
};
const URL_TABLEUR = "https://docs.google.com/spreadsheets/";
const TYPES_PAGE = ["lignee", "quadrillee", "blanche", "colonnes", "finance", "todo", "agenda", "web", "tableur"];
const TYPES_WEB = ["web", "tableur"];
const TYPES_LIBELLES = [["lignee", "Lignée"], ["quadrillee", "Quadrillée"], ["blanche", "Blanche (dessin libre)"],
  ["colonnes", "En colonnes…"], ["finance", "Finance (suivi d'un compte)"], ["todo", "To Do (local)"],
  ["web:todo", "To Do (google)"], ["agenda", "Agenda (local)"], ["web:agenda", "Agenda (google)"],
  ["web:contacts", "Contact (Google)"], ["tableur", "Tableur (Google Sheets)…"]];

/* ===================================================================== outils */

function nouvelId() {
  const o = new Uint8Array(6);
  crypto.getRandomValues(o);
  return Array.from(o, (b) => b.toString(16).padStart(2, "0")).join("");
}
const deux = (n) => String(n).padStart(2, "0");
function iso(d) { return `${d.getFullYear()}-${deux(d.getMonth() + 1)}-${deux(d.getDate())}`; }
function dateIso(t) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t || "");
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
function ajouterJours(d, n) { const r = new Date(d.getFullYear(), d.getMonth(), d.getDate()); r.setDate(r.getDate() + n); return r; }
function aujourdhui() { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function jourSemaine(d) { return (d.getDay() + 6) % 7; }   // lundi = 0, comme datetime.weekday()
function maintenantIso() {
  const d = new Date();
  return `${iso(d)}T${deux(d.getHours())}:${deux(d.getMinutes())}:${deux(d.getSeconds())}`;
}
function dateCourte(t) { const d = dateIso(t); return d ? `${d.getDate()} ${MOIS_ABR[d.getMonth()]} ${d.getFullYear()}` : ""; }
function dateLongue(d) { return `${JOURS[jourSemaine(d)]} ${d.getDate()} ${MOIS[d.getMonth()]}`; }
function quand(texte) {
  const d = new Date(texte);
  if (isNaN(d)) return "";
  const t = iso(d) === iso(new Date()) ? "aujourd'hui" : `le ${d.getDate()} ${MOIS_ABR[d.getMonth()]}`;
  return `${t} à ${d.getHours()} h ${deux(d.getMinutes())}`;
}
function sansAccents(t) { return String(t).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }
function rgb(c) { const n = parseInt(String(c).replace("#", ""), 16) || 0; return [n >> 16 & 255, n >> 8 & 255, n & 255]; }
/** Couleur entre c1 et c2 (t = 0 : c1, t = 1 : c2), comme melanger() du PC. */
function melanger(c1, c2, t) {
  const a = rgb(c1), b = rgb(c2);
  return "#" + a.map((x, i) => Math.round(x + (b[i] - x) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
}
function fonce(c, t) { return melanger(c, "#000000", t === undefined ? 0.3 : t); }
function compter(texte, car) { let n = 0; for (const c of texte) if (c === car) n++; return n; }
function clone(v) { return v === undefined ? undefined : JSON.parse(JSON.stringify(v)); }
/** JSON aux clés triées : comparaison de contenu indépendante de l'ordre des clés. */
function canon(v) {
  if (Array.isArray(v)) return "[" + v.map(canon).join(",") + "]";
  if (v && typeof v === "object") {
    return "{" + Object.keys(v).sort().filter((k) => v[k] !== undefined)
      .map((k) => JSON.stringify(k) + ":" + canon(v[k])).join(",") + "}";
  }
  return JSON.stringify(v === undefined ? null : v);
}

/** 1 330,55 (ou « 1 330,55 $ ») ; négatif entre parenthèses, comme un relevé de compte (comme sur le PC). */
function montantTexte(v, symbole) {
  const [ent, dec] = Math.abs(v).toFixed(2).split(".");
  const texte = ent.replace(/\B(?=(\d{3})+(?!\d))/g, " ") + "," + dec + (symbole ? " $" : "");
  return v < 0 ? `(${texte})` : texte;
}
/** « 1 234,56 », « 1234.56 », « (12,00) » → nombre arrondi au cent ; "" → 0 ; illisible → null. */
function lireMontant(t) {
  let s = String(t == null ? "" : t).replace(/[\s  $]/g, "");
  let negatif = false;
  if (/^\(.*\)$/.test(s)) { negatif = true; s = s.slice(1, -1); }
  if (s.startsWith("-") || s.startsWith("−")) { negatif = !negatif; s = s.slice(1); }
  if (s.includes(",") && s.includes(".")) s = s.replace(/\./g, "");
  s = s.replace(",", ".");
  if (s === "") return 0;
  if (!/^\d*\.?\d*$/.test(s) || s === ".") return null;
  const n = Math.round(Number(s) * 100) / 100;
  return negatif ? -n : n;
}

/* ----- jours fériés du Québec (mêmes règles que le PC) ----- */
function paques(a) {
  const n = a % 19, b = Math.floor(a / 100), c = a % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3);
  const h_ = (19 * n + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4;
  const l = (32 + 2 * e + 2 * i - h_ - k) % 7, m = Math.floor((n + 11 * h_ + 22 * l) / 451);
  const mois = Math.floor((h_ + l - 7 * m + 114) / 31), jour = (h_ + l - 7 * m + 114) % 31 + 1;
  return new Date(a, mois - 1, jour);
}
function niemeLundi(a, mois, n) {
  const d = new Date(a, mois - 1, 1);
  return ajouterJours(d, (7 - jourSemaine(d)) % 7 + 7 * (n - 1));
}
function feriesQuebec(a) {
  const p = paques(a), veille = new Date(a, 4, 24);
  return new Map([
    [iso(new Date(a, 0, 1)), "Jour de l'An"], [iso(ajouterJours(p, -2)), "Vendredi saint"],
    [iso(ajouterJours(p, 1)), "Lundi de Pâques"], [iso(ajouterJours(veille, -jourSemaine(veille))), "Journée des patriotes"],
    [iso(new Date(a, 5, 24)), "Fête nationale"], [iso(new Date(a, 6, 1)), "Fête du Canada"],
    [iso(niemeLundi(a, 9, 1)), "Fête du Travail"], [iso(niemeLundi(a, 10, 2)), "Action de grâce"],
    [iso(new Date(a, 11, 25)), "Noël"],
  ].sort((x, y) => (x[0] < y[0] ? -1 : 1)));
}
const _feries = new Map();
function ferie(d) {
  if (!etat.cahier || etat.cahier.feries === false) return null;
  const a = d.getFullYear();
  if (!_feries.has(a)) _feries.set(a, feriesQuebec(a));
  return _feries.get(a).get(iso(d)) || null;
}
function debutSemaine(jour, lundi) {
  const premier = lundi ? 0 : 6;
  return ajouterJours(jour, -((jourSemaine(jour) - premier + 7) % 7));
}

/* ----- construction du DOM ----- */
function h(tag, props, ...enfants) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(props || {})) {
    if (v == null || v === false) continue;
    if (k === "class") e.className = v;
    else if (k === "style" && typeof v === "object") {
      for (const [p, val] of Object.entries(v)) if (val != null) {
        if (p.startsWith("--")) e.style.setProperty(p, val); else e.style[p] = val;
      }
    } else if (k.startsWith("on") && typeof v === "function") e.addEventListener(k.slice(2), v);
    else if (k === "value" || k === "checked" || k === "disabled" || k === "selected") e[k] = v;
    else e.setAttribute(k, v === true ? "" : v);
  }
  for (const c of enfants.flat(Infinity)) if (c != null && c !== false) e.append(c instanceof Node ? c : String(c));
  return e;
}
function svg(tag, attrs, ...enfants) {
  const e = document.createElementNS("http://www.w3.org/2000/svg", tag);
  for (const [k, v] of Object.entries(attrs || {})) if (v != null) e.setAttribute(k, v);
  for (const c of enfants.flat()) if (c != null && c !== false) e.append(c instanceof Node ? c : String(c));
  return e;
}

/* ===================================================================== stockage local (IndexedDB) */

const BD = {
  _db: null,
  _memoire: new Map(),
  ouvrir() {
    if (!this._db) {
      this._db = new Promise((ok, ko) => {
        const r = indexedDB.open(CONFIG.base, 1);
        r.onupgradeneeded = () => r.result.createObjectStore("kv");
        r.onsuccess = () => ok(r.result);
        r.onerror = () => ko(r.error);
      });
    }
    return this._db;
  },
  async lire(cle) {
    try {
      const db = await this.ouvrir();
      return await new Promise((ok, ko) => {
        const r = db.transaction("kv").objectStore("kv").get(cle);
        r.onsuccess = () => ok(r.result);
        r.onerror = () => ko(r.error);
      });
    } catch (e) {
      return this._memoire.get(cle);
    }
  },
  async ecrire(cle, val) {
    this._memoire.set(cle, val);
    try {
      const db = await this.ouvrir();
      await new Promise((ok, ko) => {
        const t = db.transaction("kv", "readwrite");
        t.objectStore("kv").put(val, cle);
        t.oncomplete = () => ok();
        t.onerror = () => ko(t.error);
      });
    } catch (e) {
      if (!BD._avertie) { BD._avertie = true; avis("Stockage de l'iPhone indisponible : les modifications restent en mémoire"); }
    }
  },
};

/* ===================================================================== état */

function vierge() {
  return {cahier: null, base: null, etag: null, modifie: false, dossier: null, synchro: null, compte: null};
}
const etat = Object.assign({mode: "reel"}, vierge());
let generation = 0;           // augmente à chaque modification locale (envoi pendant une saisie)
let jetons = null;            // {acces, expire, rafraichir}
const sync = {enCours: false, erreur: null, besoinConnexion: false, horsLigne: !navigator.onLine, relancer: false,
  attenteSaisie: false, auDemarrage: true, verrouSignale: false};
const semaines = new Map();   // page Agenda -> premier jour de la semaine affichée (comme sur le PC, non enregistré)
let anneeCalendrier = null;
const ABSENCE_COUVERTURE = 2 * 60 * 1000;  // retour dans l'app après 2 min : la couverture s'affiche de nouveau
let apresCouverture = null;              // écran où reprendre après la couverture (retour dans l'app)
let cacheDepuis = 0;

async function chargerEtat() {
  etat.mode = (await BD.lire("mode")) === "demo" ? "demo" : "reel";
  Object.assign(etat, vierge(), (await BD.lire("etat-" + etat.mode)) || {});
  jetons = (await BD.lire("jetons")) || null;
}
let minuterieSauver = null;
function sauverLocal(immediat) {
  clearTimeout(minuterieSauver);
  const ecrire = () => BD.ecrire("etat-" + etat.mode, {cahier: etat.cahier, base: etat.base, etag: etat.etag,
    modifie: etat.modifie, dossier: etat.dossier, synchro: etat.synchro, compte: etat.compte});
  if (immediat) return ecrire();
  minuterieSauver = setTimeout(ecrire, 250);
  return Promise.resolve();
}
function peutModifier() { return !!etat.cahier && etat.cahier.version === VERSION_DONNEES; }
/** Après chaque modification : copie locale, puis envoi dans OneDrive quelques secondes plus tard. */
function modifier() {
  if (!peutModifier()) return;
  etat.modifie = true;
  generation++;
  sauverLocal();
  planifierSynchro(4000);
  majStatut();
}

/* ===================================================================== structure (comme sur le PC) */

function nouvellePage(modele, titre) {
  modele = modele || {type: "lignee"};
  const t = TYPES_PAGE.includes(modele.type) ? modele.type : "lignee";
  const p = {id: nouvelId(), type: t, cree: iso(aujourdhui())};
  if (titre !== undefined && titre !== null) p.note = {id: nouvelId(), titre};
  if (t === "lignee" || t === "quadrillee") Object.assign(p, {texte: "", images: []});
  else if (t === "blanche") Object.assign(p, {traits: [], images_libres: []});
  else if (t === "finance") Object.assign(p, {compte: {solde_initial: 0.0}, operations: [], modeles: []});
  else if (t === "todo") p.taches = [];
  else if (t === "agenda") p.jours = {};
  else if (t === "web") {
    const service = SERVICES[modele.service] ? modele.service : "page";
    Object.assign(p, {service, url: modele.url || SERVICES[service][1]});
  } else if (t === "tableur") p.url = modele.url || URL_TABLEUR;
  else {
    Object.assign(p, {colonnes: (modele.colonnes || ["Colonne 1", "Colonne 2"]).slice(), cellules: [], images_libres: []});
    const largeurs = largeursColonnes(p.colonnes.length, modele.largeurs);
    if (largeurs) p.largeurs = largeurs;
  }
  return p;
}
function modeleDe(p) {
  if (p.type === "colonnes") {
    const modele = {type: "colonnes", colonnes: (p.colonnes || ["Nom", "Tél."]).slice()};
    if (p.largeurs) modele.largeurs = p.largeurs.slice();
    return modele;
  }
  if (p.type === "web") return {type: "web", service: p.service || "page"};
  return {type: TYPES_PAGE.includes(p.type) ? p.type : "lignee"};
}
/** Largeur des colonnes en % (la dernière prend le reste) ; null : colonnes égales (comme largeurs_colonnes du PC). */
function largeursColonnes(n, largeurs) {
  if (!Array.isArray(largeurs) || largeurs.length !== n || n < 1) return null;
  const v = largeurs.map(Number);
  if (v.some((x) => !(x > 0))) return null;
  const total = v.reduce((a, b) => a + b, 0);
  const pct = v.slice(0, -1).map((x) => Math.max(1, Math.round(x * 100 / total)));
  const reste = 100 - pct.reduce((a, b) => a + b, 0);
  return reste >= 1 ? [...pct, reste] : null;
}
/** Police et taille d'écriture propres à la page (Paramètres de la page sur le PC). */
function policePage(p) {
  const r = p && typeof p.police === "object" && p.police ? p.police : {};
  const taille = Number(r.taille) || 1;
  return {famille: String(r.famille || "").trim() || null, taille: Math.min(1, Math.max(0.5, taille))};
}
function stylePolice(p, base) {
  const {famille, taille} = policePage(p);
  return {fontFamily: famille ? `"${famille.replace(/"/g, "")}", -apple-system, "Segoe UI", sans-serif` : null,
    fontSize: taille !== 1 ? `${Math.round(base * taille * 10) / 10}px` : null};
}
function nouvelleTache(texte) { return {id: nouvelId(), texte: texte || "", fait: false, echeance: null}; }
function nouvelleOperation(date) {
  return {date: date || iso(aujourdhui()), ref: "", tiers: "", debit: 0.0, credit: 0.0, pointe: false};
}

/** [[début, fin]] des notes de la section (fin exclue). */
function notesSection(s) {
  const notes = [];
  let debut = null;
  s.pages.forEach((p, k) => {
    if (p.type === "calendrier" || p.note) {
      if (debut !== null) notes.push([debut, k]);
      debut = p.note ? k : null;
    }
  });
  if (debut !== null) notes.push([debut, s.pages.length]);
  return notes;
}
function finNote(s, debut) { const n = notesSection(s).find((x) => x[0] === debut); return n ? n[1] : debut + 1; }
function debutNote(s, pi) {
  for (let k = Math.min(pi, s.pages.length - 1); k >= 0; k--) {
    if (s.pages[k].type === "calendrier") return null;
    if (s.pages[k].note) return k;
  }
  return null;
}
function sujetDe(s, pi) {
  if (!s.sujets) return null;
  for (let k = Math.min(pi, s.pages.length - 1); k >= 0; k--) {
    if (s.pages[k].type === "calendrier") return null;
    if (s.pages[k].sujet) return k;
  }
  return null;
}
function finSujet(s, debut) {
  for (let k = debut + 1; k < s.pages.length; k++) if (s.pages[k].sujet || s.pages[k].type === "calendrier") return k;
  return s.pages.length;
}
/** Lignes de la liste des notes : {genre, debut, titre, pages, dansSujet, notes}. */
function entreesSection(s) {
  const notes = notesSection(s);
  const entrees = [];
  s.pages.forEach((p, k) => {
    if (p.type === "calendrier") entrees.push({genre: "calendrier", debut: k, titre: "Calendrier", pages: 1, dansSujet: false, notes: 0});
  });
  for (const [debut, fin] of notes) {
    const p = s.pages[debut];
    if (p.sujet && s.sujets) {
      const fs = finSujet(s, debut);
      entrees.push({genre: "sujet", debut, titre: p.sujet.titre || "", pages: fs - debut, dansSujet: false,
        notes: notes.filter(([d]) => debut <= d && d < fs).length});
    }
    entrees.push({genre: "note", debut, titre: p.note.titre || "", pages: fin - debut, dansSujet: sujetDe(s, debut) !== null, notes: 0});
  }
  return entrees;
}
function pageVide(p) {
  const t = p.type;
  if (t === "lignee" || t === "quadrillee") return !p.texte && !(p.images || []).length;
  if (t === "blanche") return !(p.traits || []).length && !(p.images_libres || []).length;
  if (t === "colonnes") return !(p.images_libres || []).length && !(p.cellules || []).some((r) => r.some((c) => String(c).trim()));
  if (t === "finance") return !(p.operations || []).length;
  if (t === "todo") return !(p.taches || []).length;
  if (t === "agenda") return !Object.values(p.jours || {}).some((l) => l.some((x) => String(x).trim()));
  return TYPES_WEB.includes(t);
}
/** Ajoute une page à la note [début, fin) : « suite » (après pi), « debut » (elle porte alors le titre et le sujet) ou « fin ». */
function ajouterPageNote(s, debut, fin, pi, modele, position) {
  const pos = position === "debut" ? debut : position === "fin" ? fin : pi + 1;
  const nouvelle = nouvellePage(modele || modeleDe(s.pages[pi]));
  s.pages.splice(pos, 0, nouvelle);
  if (pos === debut) {
    const ancienne = s.pages[debut + 1];
    for (const cle of ["note", "sujet"]) if (cle in ancienne) { nouvelle[cle] = ancienne[cle]; delete ancienne[cle]; }
  }
  return nouvelle;
}
/** Déplace la page pi dans sa note : « avant », « apres », « debut », « fin » ; nouvel indice, ou null. */
function deplacerPage(s, debut, fin, pi, cible) {
  const k = {avant: pi - 1, apres: pi + 1, debut, fin: fin - 1}[cible];
  if (!(debut <= k && k < fin) || k === pi) return null;
  const premiere = s.pages[debut];
  const marqueurs = {};
  for (const cle of ["note", "sujet"]) if (cle in premiere) { marqueurs[cle] = premiere[cle]; delete premiere[cle]; }
  const [page] = s.pages.splice(pi, 1);
  s.pages.splice(k, 0, page);
  Object.assign(s.pages[debut], marqueurs);
  return k;
}
/** Supprime les pages a à b - 1 ; le titre de la note et le sujet passent à la page suivante qui les continue. */
function supprimerPages(s, a, b) {
  const pages = s.pages;
  const supprimees = pages.slice(a, b);
  const premiere = pages[a];
  if (b < pages.length && pages[b].type !== "calendrier") {
    const suivante = pages[b];
    if (premiere.note && !suivante.note) suivante.note = premiere.note;
    if (premiere.sujet && !suivante.sujet && s.sujets) {
      suivante.sujet = premiere.sujet;
      if (!suivante.note) suivante.note = {id: nouvelId(), titre: ""};
    }
  }
  pages.splice(a, b - a);
  return supprimees;
}
function poserCellule(p, i, k, valeur) {
  const cases = p.cellules, n = p.colonnes.length;
  while (cases.length <= i) cases.push(new Array(n).fill(""));
  const rang = cases[i];
  while (rang.length < n) rang.push("");
  rang[k] = valeur;
  while (cases.length && !cases[cases.length - 1].some((v) => String(v).trim())) cases.pop();
}
/** Saisie dans le texte d'une page : les images (caractères OBJ) effacées sont retirées de la page ; on ne peut pas
 *  en ajouter en tapant. Retourne le texte gardé. */
function appliquerSaisieTexte(p, ancien, nouveau) {
  let pre = 0;
  const max = Math.min(ancien.length, nouveau.length);
  while (pre < max && ancien[pre] === nouveau[pre]) pre++;
  let suf = 0;
  while (suf < max - pre && ancien[ancien.length - 1 - suf] === nouveau[nouveau.length - 1 - suf]) suf++;
  const retire = ancien.slice(pre, ancien.length - suf);
  let ajoute = nouveau.slice(pre, nouveau.length - suf);
  const n = compter(retire, OBJ);
  if (n) (p.images || (p.images = [])).splice(compter(ancien.slice(0, pre), OBJ), n);
  if (ajoute.includes(OBJ)) ajoute = ajoute.split(OBJ).join("");
  p.texte = nouveau.slice(0, pre) + ajoute + nouveau.slice(nouveau.length - suf);
  return p.texte;
}
function soldes(p) {
  const triees = p.operations.slice().sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  let solde = Number(p.compte.solde_initial) || 0;
  const liste = triees.map((o) => (solde = Math.round((solde + (+o.credit || 0) - (+o.debit || 0)) * 100) / 100));
  return {triees, soldes: liste};
}
function trierModeles(p) { p.modeles.sort((a, b) => (sansAccents(a.libelle) < sansAccents(b.libelle) ? -1 : sansAccents(a.libelle) > sansAccents(b.libelle) ? 1 : 0)); }
function libelleType(p) {
  if (p.type === "web") return (SERVICES[p.service] || SERVICES.page)[0];
  if (p.type === "calendrier") return "Calendrier";
  const t = TYPES_LIBELLES.find(([v]) => v === p.type);
  return t ? t[1].replace(/…$/, "").replace(/ \(.*\)$/, "") : p.type;
}
function trouverSection(sid) {
  const si = etat.cahier.sections.findIndex((s) => s.id === sid);
  return si < 0 ? null : etat.cahier.sections[si];
}

/* ===================================================================== fusion des versions */

function structure(c) {
  return canon(c.sections.map((s) => [s.id, s.titre, s.couleur, s.type, !!s.sujets, s.modele_page || null,
    s.pages.map((p) => [p.id, p.note || null, p.sujet || null])]));
}
function contenu(p) { const r = Object.assign({}, p); delete r.note; delete r.sujet; return canon(r); }
function pagesParId(c) { const m = new Map(); for (const s of c.sections) for (const p of s.pages) m.set(p.id, p); return m; }
function remplacerContenu(cible, source) {
  for (const k of Object.keys(cible)) if (k !== "note" && k !== "sujet") delete cible[k];
  for (const [k, v] of Object.entries(source)) if (k !== "note" && k !== "sujet") cible[k] = clone(v);
}
/** Fusion à trois : base (dernière version reçue), locale (iPhone), distante (OneDrive). Les pages modifiées d'un seul
 *  côté sont gardées ; si la structure (sections, notes, ordre des pages) n'a changé que d'un côté au plus et qu'aucune
 *  page n'a été modifiée des deux côtés, {cahier: fusion, copie: null}. Sinon {cahier: distante, copie: locale}. */
function fusionner(base, locale, distante) {
  const echec = {cahier: distante, copie: locale};
  if (!base || distante.version !== VERSION_DONNEES || !Array.isArray(distante.sections)) return echec;
  const sB = structure(base), sL = structure(locale), sR = structure(distante);
  let resultat, autre;
  if (sL === sB) { resultat = clone(distante); autre = locale; } else if (sR === sB) { resultat = clone(locale); autre = distante; } else return echec;
  const pB = pagesParId(base), pRes = pagesParId(resultat);
  for (const [id, p] of pagesParId(autre)) {
    const b = pB.get(id), cible = pRes.get(id);
    if (b && contenu(p) === contenu(b)) continue;
    if (!b || !cible) return echec;
    if (contenu(cible) === contenu(p)) continue;
    if (contenu(cible) !== contenu(b)) return echec;   // la même page modifiée des deux côtés
    remplacerContenu(cible, p);
  }
  const cles = new Set([...Object.keys(base), ...Object.keys(locale), ...Object.keys(distante)]);
  for (const k of cles) {
    if (k === "sections" || k === "modifie_par" || k === "version") continue;
    const b = canon(base[k]), l = canon(locale[k]), r = canon(distante[k]);
    if (l === r) continue;
    let v;
    if (l === b) v = distante[k]; else if (r === b) v = locale[k]; else return echec;
    if (v === undefined) delete resultat[k]; else resultat[k] = clone(v);
  }
  resultat.version = distante.version;
  if (distante.modifie_par) resultat.modifie_par = clone(distante.modifie_par);
  return {cahier: resultat, copie: null};
}

/* ===================================================================== connexion Microsoft (PKCE) */

class ErreurConnexion extends Error { constructor() { super("Connexion Microsoft nécessaire"); } }
class ErreurGraph extends Error {
  constructor(statut, texte) { super(`OneDrive a répondu ${statut}${texte ? " : " + texte : ""}`); this.statut = statut; }
}
function adresseRetour() { return CONFIG.redirection || location.origin + location.pathname.replace(/index\.html$/, ""); }
function b64url(octets) {
  let s = "";
  for (const b of new Uint8Array(octets)) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function aleatoire(n) { const o = new Uint8Array(n); crypto.getRandomValues(o); return b64url(o); }
function lireLocal(cle) { try { return JSON.parse(localStorage.getItem(cle)); } catch (e) { return null; } }
function ecrireLocal(cle, v) { try { if (v == null) localStorage.removeItem(cle); else localStorage.setItem(cle, JSON.stringify(v)); } catch (e) { /* stockage refusé */ } }

async function connecter(silencieux) {
  if (!CONFIG.clientId) { avis("L'app n'est pas encore reliée à Microsoft (voir le guide)"); return; }
  if (!navigator.onLine) { avis("Connexion Internet nécessaire"); return; }
  if (!window.crypto || !crypto.subtle) { avis("Connexion possible seulement depuis l'adresse https de l'app"); return; }
  await sauverLocal(true);
  const verificateur = aleatoire(48), etatOauth = aleatoire(18);
  const defi = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verificateur)));
  ecrireLocal("agenda-pkce", {verificateur, etat: etatOauth, silencieux: !!silencieux, retour: location.hash, quand: Date.now()});
  if (silencieux) ecrireLocal("agenda-silencieux", Date.now());
  const p = new URLSearchParams({client_id: CONFIG.clientId, response_type: "code", redirect_uri: adresseRetour(),
    response_mode: "query", scope: PORTEES, state: etatOauth, code_challenge: defi, code_challenge_method: "S256",
    prompt: silencieux ? "none" : "select_account"});
  if (etat.compte && etat.compte.courriel) p.set("login_hint", etat.compte.courriel);
  location.assign(`${CONNEXION}${CONFIG.autorite}/oauth2/v2.0/authorize?${p}`);
}
async function garderJetons(d) {
  jetons = {acces: d.access_token, expire: Date.now() + (Number(d.expires_in) || 3600) * 1000 - 120000,
    rafraichir: d.refresh_token || (jetons && jetons.rafraichir) || null};
  await BD.ecrire("jetons", jetons);
}
async function demanderJetons(params) {
  const r = await fetch(`${CONNEXION}${CONFIG.autorite}/oauth2/v2.0/token`, {method: "POST",
    headers: {"Content-Type": "application/x-www-form-urlencoded"},
    body: new URLSearchParams(Object.assign({client_id: CONFIG.clientId, scope: PORTEES}, params))});
  const d = await r.json().catch(() => ({}));
  return {ok: r.ok, statut: r.status, d};
}
/** Retour de la page de connexion de Microsoft (?code=… ou ?error=…). */
async function retourConnexion() {
  const q = new URLSearchParams(location.search);
  if (!q.has("code") && !q.has("error")) return;
  const attente = lireLocal("agenda-pkce");
  ecrireLocal("agenda-pkce", null);
  history.replaceState(null, "", adresseRetour() + ((attente && attente.retour) || ""));
  if (!attente || q.get("state") !== attente.etat) { sync.erreur = "Réponse de connexion inattendue : reconnectez-vous."; return; }
  if (q.has("error")) {
    if (attente.silencieux) { sync.besoinConnexion = true; return; }
    sync.erreur = "Connexion refusée : " + (q.get("error_description") || q.get("error")).split(/[\r\n]/)[0];
    return;
  }
  const r = await demanderJetons({grant_type: "authorization_code", code: q.get("code"), redirect_uri: adresseRetour(),
    code_verifier: attente.verificateur});
  if (!r.ok) { sync.erreur = "Connexion impossible : " + String(r.d.error_description || r.d.error || r.statut).split(/[\r\n]/)[0]; return; }
  await garderJetons(r.d);
  sync.besoinConnexion = false;
  if (etat.mode !== "reel") await changerMode("reel");
  try {
    const moi = await graphJson("/me?$select=displayName,userPrincipalName,mail");
    etat.compte = {nom: moi.displayName || "", courriel: moi.userPrincipalName || moi.mail || ""};
    await sauverLocal(true);
  } catch (e) { /* le nom du compte n'est qu'informatif */ }
}
let rafraichissement = null;
async function jeton(force) {
  if (!force && jetons && jetons.acces && Date.now() < jetons.expire) return jetons.acces;
  if (!jetons || !jetons.rafraichir) throw new ErreurConnexion();
  if (!rafraichissement) {
    rafraichissement = (async () => {
      const r = await demanderJetons({grant_type: "refresh_token", refresh_token: jetons.rafraichir});
      if (!r.ok) {
        if (r.statut === 400 || r.statut === 401) { jetons = null; await BD.ecrire("jetons", null); throw new ErreurConnexion(); }
        throw new Error("Service de connexion indisponible (" + r.statut + ")");
      }
      await garderJetons(r.d);
      return jetons.acces;
    })().finally(() => { rafraichissement = null; });
  }
  return rafraichissement;
}
async function graph(chemin, options, essai) {
  options = options || {};
  const t = await jeton(!!essai);
  const r = await fetch(chemin.startsWith("http") ? chemin : GRAPH + chemin,
    Object.assign({}, options, {headers: Object.assign({}, options.headers || {}, {Authorization: "Bearer " + t})}));
  if (r.status === 401) {
    if (!essai) return graph(chemin, options, 1);
    throw new ErreurConnexion();
  }
  return r;
}
async function texteErreur(r) { try { const d = await r.json(); return (d.error && d.error.message) || ""; } catch (e) { return ""; } }
async function graphJson(chemin) {
  const r = await graph(chemin);
  if (!r.ok) throw new ErreurGraph(r.status, await texteErreur(r));
  return r.json();
}
function segment(s) { return encodeURIComponent(s).replace(/'/g, "%27"); }
/** Adresse Graph d'un fichier du dossier de l'agenda : /me/drive/items/{dossier}:/a/b: */
function cheminElement(...segments) { return `/me/drive/items/${etat.dossier.id}:/` + segments.map(segment).join("/") + ":"; }

/* ===================================================================== synchronisation OneDrive */

let minuterieSynchro = null;
function planifierSynchro(delai) {
  clearTimeout(minuterieSynchro);
  minuterieSynchro = setTimeout(() => synchroniser(), delai || 0);
}
function enSaisie() {
  const a = document.activeElement;
  return !!a && ["INPUT", "TEXTAREA", "SELECT"].includes(a.tagName) && !!a.closest("#app");
}
async function synchroniser(manuel) {
  if (etat.mode !== "reel" || !etat.dossier || !CONFIG.clientId) return;
  if (sync.enCours) { sync.relancer = true; return; }
  if (!navigator.onLine) { sync.horsLigne = true; majStatut(); return; }
  if (!jetons) { sync.besoinConnexion = true; majStatut(); return; }
  const auDemarrage = sync.auDemarrage;
  sync.auDemarrage = false;
  sync.enCours = true;
  sync.erreur = null;
  majStatut();
  try {
    await synchroniserUneFois(0);
    sync.besoinConnexion = false;
    sync.horsLigne = false;
    if (manuel) avis(etat.modifie ? "Modifications en attente d'envoi" : "Agenda à jour");
  } catch (e) {
    if (e instanceof ErreurConnexion) {
      sync.besoinConnexion = true;
      // Une fois par jour environ, la connexion doit être renouvelée : passage éclair par Microsoft, sans rien taper.
      const dernier = lireLocal("agenda-silencieux") || 0;
      if (auDemarrage && Date.now() - dernier > 10 * 60 * 1000 && !enSaisie() && !feuilleOuverte()) { connecter(true); return; }
    } else if (e instanceof TypeError) sync.horsLigne = true;   // réseau indisponible
    else sync.erreur = e.message || String(e);
  } finally {
    sync.enCours = false;
    majStatut();
    if (sync.relancer) { sync.relancer = false; planifierSynchro(500); }
  }
}
async function lireMeta(nom) {
  const r = await graph(cheminElement(nom));
  if (r.status === 404) return null;
  if (!r.ok) throw new ErreurGraph(r.status, await texteErreur(r));
  return r.json();
}
async function telecharger(meta) {
  const r = await fetch(meta["@microsoft.graph.downloadUrl"], {cache: "no-store"});
  if (!r.ok) throw new Error("Téléchargement de l'agenda impossible (" + r.status + ")");
  return r.text();
}
/** La version reçue de OneDrive devient la copie locale. */
function adopter(texte, etag) {
  const c = JSON.parse(texte);
  if (!c || !Array.isArray(c.sections)) throw new Error("cahier.json illisible dans OneDrive");
  etat.cahier = c;
  etat.base = texte;
  etat.etag = etag;
  etat.modifie = false;
  _feries.clear();
  sauverLocal();
  rendre();
}
async function synchroniserUneFois(tour) {
  const meta = await lireMeta(FICHIER);
  if (!meta) throw new Error(`cahier.json est introuvable dans OneDrive › ${etat.dossier.chemin}`);
  const occupe = enSaisie() || feuilleOuverte();
  if (!etat.modifie || !etat.cahier) {
    if (meta.eTag !== etat.etag || !etat.cahier) {
      if (occupe && etat.cahier) { sync.attenteSaisie = true; return; }   // pas de changement sous les doigts
      const gen = generation;
      const texte = await telecharger(meta);
      if (gen !== generation) return synchroniserUneFois(tour + 1);
      adopter(texte, meta.eTag);
      if (etat.synchro && etat.cahier.modifie_par && etat.cahier.modifie_par.poste !== POSTE) {
        avis(`Mis à jour depuis le poste ${etat.cahier.modifie_par.poste}`);
      }
    }
    etat.synchro = new Date().toISOString();
    sauverLocal();
    return;
  }
  if (meta.eTag === etat.etag) { await envoyer(meta.eTag, tour); return; }
  // L'agenda a aussi changé dans OneDrive (sur le PC) : fusion.
  if (occupe) { sync.attenteSaisie = true; return; }
  const gen = generation;
  const texte = await telecharger(meta);
  if (gen !== generation || enSaisie() || feuilleOuverte()) { sync.attenteSaisie = true; return; }
  const distant = JSON.parse(texte);
  const auteur = (distant.modifie_par && distant.modifie_par.poste) || "un autre poste";
  const res = fusionner(etat.base ? JSON.parse(etat.base) : null, etat.cahier, distant);
  if (res.copie) {
    const nom = await enregistrerCopieConflit(res.copie);
    adopter(texte, meta.eTag);
    etat.synchro = new Date().toISOString();
    sauverLocal();
    message("Agenda modifié sur un autre poste",
      `L'agenda a été modifié sur le poste « ${auteur} » pendant que vous le modifiiez sur l'iPhone.\n\n` +
      `La version de l'iPhone a été gardée dans OneDrive › ${etat.dossier.chemin} › ${nom}.\n\n` +
      `L'iPhone affiche maintenant la version du poste « ${auteur} ».`);
    return;
  }
  etat.cahier = res.cahier;
  etat.base = texte;
  etat.etag = meta.eTag;
  sauverLocal();
  rendre();
  await envoyer(meta.eTag, tour);
  avis(`Vos modifications ont été fusionnées avec celles du poste ${auteur}`);
}
async function envoyer(etag, tour) {
  const gen = generation;
  etat.cahier.modifie_par = {poste: POSTE, utilisateur: (etat.compte && etat.compte.nom) || "", date: maintenantIso()};
  const texte = JSON.stringify(etat.cahier);
  const r = await graph(cheminElement(FICHIER) + "/content", {method: "PUT",
    headers: {"Content-Type": "application/json", "If-Match": etag}, body: texte});
  if (r.status === 412 || r.status === 409) {
    if (tour < 2) return synchroniserUneFois(tour + 1);
    throw new Error("L'agenda change sans arrêt dans OneDrive : nouvel essai dans un moment.");
  }
  if (!r.ok) throw new ErreurGraph(r.status, await texteErreur(r));
  const meta = await r.json();
  etat.etag = meta.eTag;
  etat.base = texte;
  etat.modifie = generation !== gen;
  etat.synchro = new Date().toISOString();
  sauverLocal();
  if (etat.modifie) planifierSynchro(1500);
  verifierVerrou();
}
async function enregistrerCopieConflit(cahier) {
  const d = new Date();
  const nom = `cahier-conflit-${POSTE}-${iso(d).replace(/-/g, "")}-${deux(d.getHours())}${deux(d.getMinutes())}${deux(d.getSeconds())}.json`;
  const r = await graph(cheminElement(nom) + "/content?@microsoft.graph.conflictBehavior=rename",
    {method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify(cahier)});
  if (!r.ok) throw new ErreurGraph(r.status, await texteErreur(r));
  return (await r.json()).name || nom;
}
/** Verrou du PC (verrou.json) : seulement pour dire que l'agenda est aussi ouvert sur un poste, qui se mettra à jour. */
async function verifierVerrou() {
  if (sync.verrouSignale) return;
  try {
    const meta = await lireMeta("verrou.json");
    if (!meta) return;
    const v = JSON.parse(await telecharger(meta));
    if (v && v.poste && Date.now() / 1000 - Number(v.battement || 0) < DUREE_VERROU) {
      sync.verrouSignale = true;
      avis(`L'agenda est aussi ouvert sur le poste « ${v.poste} » : il s'y mettra à jour tout seul`, 5000);
    }
  } catch (e) { /* information seulement */ }
}

/* ----- fichiers du dossier (pièces jointes, images) ----- */
async function adresseTelechargement(segments) {
  const r = await graph(cheminElement(...segments));
  if (!r.ok) return null;
  return (await r.json())["@microsoft.graph.downloadUrl"] || null;
}
async function nomLibre(pid, nom) {
  const point = nom.lastIndexOf(".");
  const base = point > 0 ? nom.slice(0, point) : nom, ext = point > 0 ? nom.slice(point) : "";
  for (let k = 1; k < 100; k++) {
    const essai = k === 1 ? nom : `${base} (${k})${ext}`;
    const r = await graph(cheminElement("documents", pid, essai));
    if (r.status === 404) return essai;
    if (!r.ok) throw new ErreurGraph(r.status, await texteErreur(r));
  }
  throw new Error("Trop de fichiers du même nom");
}
async function envoyerPiece(pid, fichier) {
  const nom = await nomLibre(pid, (fichier.name || "fichier").replace(/[\\/:*?"<>|]/g, "_"));
  const chemin = cheminElement("documents", pid, nom);
  if (fichier.size <= 4 * 1024 * 1024) {
    const r = await graph(chemin + "/content?@microsoft.graph.conflictBehavior=fail",
      {method: "PUT", headers: {"Content-Type": fichier.type || "application/octet-stream"}, body: fichier});
    if (!r.ok) throw new ErreurGraph(r.status, await texteErreur(r));
    return nom;
  }
  const r = await graph(chemin + "/createUploadSession", {method: "POST", headers: {"Content-Type": "application/json"},
    body: JSON.stringify({item: {"@microsoft.graph.conflictBehavior": "fail"}})});
  if (!r.ok) throw new ErreurGraph(r.status, await texteErreur(r));
  const {uploadUrl} = await r.json();
  const BLOC = 320 * 1024 * 10;
  for (let debut = 0; debut < fichier.size; debut += BLOC) {
    const fin = Math.min(debut + BLOC, fichier.size);
    const rr = await fetch(uploadUrl, {method: "PUT", headers: {"Content-Range": `bytes ${debut}-${fin - 1}/${fichier.size}`},
      body: fichier.slice(debut, fin)});
    if (!rr.ok) throw new Error("Envoi du fichier interrompu (" + rr.status + ")");
  }
  return nom;
}
async function supprimerDansOneDrive(...segments) {
  if (etat.mode !== "reel" || !etat.dossier || !jetons || !navigator.onLine) return false;
  try { const r = await graph(cheminElement(...segments), {method: "DELETE"}); return r.ok || r.status === 404; } catch (e) { return false; }
}
const _images = new Map();
/** Image du dossier images\ de l'agenda : depuis le cache de l'iPhone, sinon depuis OneDrive (gardée pour hors ligne). */
async function urlImage(fichier) {
  if (_images.has(fichier)) return _images.get(fichier);
  const cle = new URL("__images/" + encodeURIComponent(fichier), location.href).href;
  let blob = null;
  try { const r = await (await caches.open(CACHE_IMAGES)).match(cle); if (r) blob = await r.blob(); } catch (e) { /* pas de cache */ }
  if (!blob && etat.mode === "reel" && etat.dossier && jetons && navigator.onLine) {
    try {
      const url = await adresseTelechargement(["images", fichier]);
      const r = url ? await fetch(url) : null;
      if (r && r.ok) {
        blob = await r.blob();
        try { await (await caches.open(CACHE_IMAGES)).put(cle, new Response(blob)); } catch (e) { /* cache plein */ }
      }
    } catch (e) { /* hors ligne */ }
  }
  if (!blob) return null;
  const url = URL.createObjectURL(blob);
  _images.set(fichier, url);
  return url;
}

/* ===================================================================== fenêtres (feuille du bas), avis */

let fermerCourante = null;
function feuilleOuverte() { const z = document.getElementById("feuille"); return !!z && !z.hidden; }
function ouvrirFeuille(construire) {
  if (fermerCourante) fermerCourante(null);
  return new Promise((ok) => {
    const zone = document.getElementById("feuille");
    let fermee = false;
    const fermer = (valeur) => {
      if (fermee) return;
      fermee = true;
      if (fermerCourante === fermer) { fermerCourante = null; zone.hidden = true; zone.replaceChildren(); }
      ok(valeur);
      setTimeout(apresSaisie, 60);
    };
    fermerCourante = fermer;
    const panneau = h("div", {class: "panneau", role: "dialog", "aria-modal": "true"}, h("div", {class: "poignee"}), construire(fermer));
    // Un toucher qui arrive juste à l'ouverture (double toucher) ne doit ni fermer la fenêtre ni choisir à sa place.
    const ouverte = performance.now();
    panneau.addEventListener("click", (e) => {
      if (performance.now() - ouverte < 400) { e.stopPropagation(); e.preventDefault(); }
    }, true);
    zone.replaceChildren(panneau);
    zone.hidden = false;
    zone.onclick = (e) => { if (e.target === zone && performance.now() - ouverte > 400) fermer(null); };
    const champ = panneau.querySelector("input:not([type=checkbox]):not([type=file]):not([type=date]), textarea");
    if (champ) { champ.focus(); if (champ.select) champ.select(); }
  });
}
function choisir(titre, options, texte) {
  return ouvrirFeuille((fermer) => [h("h2", {}, titre), texte ? h("p", {class: "message"}, texte) : null,
    h("div", {class: "choix"}, options.filter(Boolean).map((o) => h("button", {class: "bouton" + (o.danger ? " danger" : "") + (o.principal ? " principal" : ""),
      onclick: () => fermer(o.valeur)}, o.libelle))),
    h("div", {class: "actions"}, h("button", {class: "bouton", onclick: () => fermer(null)}, "Annuler"))]);
}
function confirmer(texte, libelle, danger) {
  return ouvrirFeuille((fermer) => [h("p", {class: "message"}, texte), h("div", {class: "actions"},
    h("button", {class: "bouton", onclick: () => fermer(false)}, "Annuler"),
    h("button", {class: "bouton " + (danger === false ? "principal" : "danger"), onclick: () => fermer(true)}, libelle || "Supprimer"))])
    .then((v) => v === true);
}
function message(titre, texte) {
  return ouvrirFeuille((fermer) => [h("h2", {}, titre), h("p", {class: "message"}, texte),
    h("div", {class: "actions"}, h("button", {class: "bouton principal", onclick: () => fermer(true)}, "OK"))]);
}
/** Petit formulaire : champs [{cle, libelle, type: text|select|date, valeur, options: [[valeur, libellé]]}]. */
function formulaire(titre, champs, libelleOk) {
  return ouvrirFeuille((fermer) => {
    const entrees = {};
    const valider = () => {
      const r = {};
      for (const [cle, e] of Object.entries(entrees)) r[cle] = e.value;
      fermer(r);
    };
    const corps = champs.map((c) => {
      let e;
      if (c.type === "select") {
        e = h("select", {}, c.options.map(([v, l]) => h("option", {value: v, selected: String(v) === String(c.valeur)}, l)));
      } else {
        e = h("input", {type: c.type || "text", value: c.valeur == null ? "" : c.valeur, autocapitalize: c.majuscules === false ? "off" : "sentences",
          inputmode: c.inputmode || null, placeholder: c.indice || null,
          onkeydown: (ev) => { if (ev.key === "Enter") { ev.preventDefault(); valider(); } }});
      }
      entrees[c.cle] = e;
      return h("label", {class: "champ"}, h("span", {}, c.libelle), e);
    });
    return [h("h2", {}, titre), corps, h("div", {class: "actions"},
      h("button", {class: "bouton", onclick: () => fermer(null)}, "Annuler"),
      h("button", {class: "bouton principal", onclick: valider}, libelleOk || "OK"))];
  });
}
async function demanderTexte(titre, libelle, valeur) {
  const r = await formulaire(titre, [{cle: "t", libelle, valeur}], "OK");
  return r ? r.t : null;
}
let minuterieAvis = null;
function avis(texte, duree) {
  const a = document.getElementById("avis");
  if (!a) return;
  a.textContent = texte;
  a.classList.add("visible");
  clearTimeout(minuterieAvis);
  minuterieAvis = setTimeout(() => a.classList.remove("visible"), duree || 3200);
}
function apresSaisie() {
  if (sync.attenteSaisie && !enSaisie() && !feuilleOuverte()) { sync.attenteSaisie = false; planifierSynchro(200); }
}

/* ===================================================================== navigation */

function route() { return location.hash.replace(/^#\/?/, "").split("/").filter((x) => x !== "").map(decodeURIComponent); }
function aller(hash) { if (location.hash === hash) rendre(); else location.hash = hash; }
function lienNote(s, debut, k) { return `#/n/${encodeURIComponent(s.id)}/${encodeURIComponent(s.pages[debut].id)}/${k || 0}`; }
let derniereRoute = null;

function statutCourant() {
  if (etat.mode === "demo") return {classe: "attente", texte: "Démo", action: () => aller("#/reglages")};
  if (!CONFIG.clientId || (!etat.dossier && !jetons)) return null;
  if (sync.enCours) return {classe: "encours", texte: "Synchro…", action: () => {}};
  if (sync.besoinConnexion || !jetons) return {classe: "erreur", texte: "Se connecter", action: () => connecter(false)};
  if (sync.horsLigne) return {classe: "attente", texte: etat.modifie ? "Hors ligne · à envoyer" : "Hors ligne",
    action: () => avis(etat.modifie ? "Hors ligne : les modifications seront envoyées au retour du réseau" : "Hors ligne : l'agenda reste consultable et modifiable")};
  if (sync.erreur) return {classe: "erreur", texte: "Erreur", action: async () => {
    const r = await choisir("Synchronisation", [{valeur: 1, libelle: "Réessayer", principal: true}], sync.erreur);
    if (r) synchroniser(true);
  }};
  if (etat.modifie) return {classe: "attente", texte: "À envoyer", action: () => synchroniser(true)};
  return {classe: "", texte: "À jour", action: () => synchroniser(true)};
}
function majStatut() {
  const st = statutCourant();
  for (const b of document.querySelectorAll(".barre .statut")) {
    if (!st) { b.hidden = true; continue; }
    b.hidden = false;
    b.className = "statut " + st.classe;
    b.querySelector(".texte").textContent = st.texte;
    b.onclick = st.action;
  }
}
function barre(o) {
  return h("header", {class: "barre", style: {"--couleur": o.couleur}},
    o.retour ? h("button", {class: "retour", "aria-label": "Retour", onclick: () => aller(o.retour)}, "‹") : null,
    h("h1", {}, o.titre),
    o.actions || null,
    h("button", {class: "statut", hidden: true}, h("span", {class: "point"}), h("span", {class: "texte"})));
}
function bandeaux() {
  const b = [];
  if (etat.cahier && etat.cahier.version !== VERSION_DONNEES) {
    b.push(h("div", {class: "bandeau"}, etat.cahier.version > VERSION_DONNEES
      ? "Cet agenda vient d'une version plus récente de l'application : il est affiché en lecture seule. Mettez l'app à jour."
      : "Cet agenda est dans un ancien format : ouvrez-le une fois sur le PC (version 1.11 ou plus) ; en attendant, il est en lecture seule."));
  }
  if (sync.erreur && etat.mode === "reel") b.push(h("div", {class: "bandeau erreur"}, sync.erreur, " ",
    h("button", {class: "lien", onclick: () => synchroniser(true)}, "Réessayer")));
  return b;
}

function rendre() {
  const app = document.getElementById("app");
  const r = route();
  let vue;
  try {
    if (r[0] === "reglages") vue = ecranReglages();
    else if (r[0] === "couverture" && etat.cahier) vue = ecranCouverture();
    else if (r[0] === "dossier" || (etat.mode === "reel" && jetons && !etat.dossier)) vue = ecranDossier();
    else if (!etat.cahier) vue = ecranBienvenue();
    else if (r[0] === "s") vue = ecranSection(r[1]);
    else if (r[0] === "n") vue = ecranNote(r[1], r[2], Number(r[3]) || 0);
    else if (r[0] === "c") vue = ecranCalendrier(r[1]);
    else vue = ecranAccueil();
  } catch (e) {
    console.error(e);
    vue = [barre({titre: "Agenda", retour: "#/"}), h("main", {class: "contenu"}, h("div", {class: "bandeau erreur"}, "Affichage impossible : " + e.message))];
  }
  const meme = location.hash === derniereRoute;
  const y = window.scrollY;
  app.replaceChildren(...vue);
  majStatut();
  window.scrollTo(0, meme ? y : 0);
  derniereRoute = location.hash;
  for (const ta of app.querySelectorAll("textarea.zone-texte")) ajusterHauteur(ta);
}

/* ===================================================================== écrans */

function ecranBienvenue() {
  const pret = !!CONFIG.clientId;
  const attente = etat.mode === "reel" && jetons && etat.dossier;
  return [barre({titre: "Agenda", actions: h("button", {class: "icone", "aria-label": "Réglages", onclick: () => aller("#/reglages")}, "⚙")}),
    ...bandeaux(),
    h("main", {class: "contenu"}, h("div", {class: "accueil-bienvenue"},
      h("img", {src: "icones/apple-touch-icon.png", alt: ""}),
      h("h2", {}, attente ? "Téléchargement de l'agenda…" : "Votre agenda sur l'iPhone"),
      attente
        ? h("p", {}, sync.enCours ? "Lecture de cahier.json dans OneDrive." : "L'agenda sera téléchargé dès que le réseau sera disponible.")
        : h("p", {}, "Connectez-vous avec le compte Microsoft de votre OneDrive : l'agenda est copié sur l'iPhone, puis il fonctionne aussi sans réseau."),
      !attente && !pret ? h("div", {class: "bandeau"}, "L'app n'est pas encore reliée à Microsoft : inscrivez-la dans Microsoft Entra, puis inscrivez son ID dans config.js (voir le guide). Adresse de retour à donner : ",
        h("b", {}, adresseRetour())) : null,
      h("div", {class: "boutons", style: {flexDirection: "column"}},
        attente ? h("button", {class: "bouton principal large", onclick: () => synchroniser(true)}, "Réessayer")
          : h("button", {class: "bouton principal large", disabled: !pret, onclick: () => connecter(false)}, "Se connecter avec Microsoft"),
        h("button", {class: "bouton large", onclick: () => ouvrirDemo(false)}, "Essayer la démonstration"))))];
}

/* ----- couverture : la même que sur le PC (cuir, spirale, couture, titre doré, photo, année, élastique) ----- */
const COUL_OR = "#D9B65E", COUL_OR_CLAIR = "#F4E3A8";
let _mesureur = null;
function largeurTexte(texte, police) {
  _mesureur = _mesureur || document.createElement("canvas").getContext("2d");
  _mesureur.font = police;
  return _mesureur.measureText(texte).width;
}
function texteDore(x, y, texte, attrs, fond) {
  const t = (dx, dy, couleur) => svg("text", Object.assign({x: x + dx, y: y + dy, fill: couleur, "text-anchor": "middle",
    "dominant-baseline": "central"}, attrs), texte);
  return [t(0.07, 0.13, fonce(fond, 0.55)), t(-0.06, -0.06, COUL_OR_CLAIR), t(0, 0, COUL_OR)];
}
function illustration(x, y, w, h) {
  const p = (points, couleur) => svg("polygon", {points: points.map(([a, b]) => `${x + w * a},${y + h * b}`).join(" "), fill: couleur});
  const sx = x + w * 0.68, sy = y + h * 0.42;
  return [svg("rect", {x, y, width: w, height: h * 0.62 + 0.02, fill: "url(#ciel)"}),
    svg("circle", {cx: sx, cy: sy, r: w * 0.085, fill: "#FFD27A"}),
    p([[0, .62], [.17, .33], [.33, .5], [.52, .27], [.74, .54], [.88, .4], [1, .5], [1, .62]], "#7E9BB7"),
    p([[0, .62], [.1, .5], [.24, .58], [.42, .44], [.6, .62]], "#5B7894"),
    svg("rect", {x, y: y + h * 0.62, width: w, height: h * 0.38, fill: "#5F90B6"}),
    [[.55, .8], [.6, .75], [.62, .72]].map(([a, b], k) => svg("line", {x1: x + w * a, x2: x + w * b,
      y1: y + h * (0.66 + k * 0.035), y2: y + h * (0.66 + k * 0.035), stroke: "#A9CBE3", "stroke-width": 0.08})),
    p([[0, .8], [.3, .72], [.55, .79], [.8, .74], [1, .77], [1, 1], [0, 1]], "#3F6B4F"),
    [[.08, .78, .09], [.15, .76, .12], [.24, .75, .08], [.84, .74, .1], [.92, .76, .07]].map(([a, b, s]) =>
      p([[a, b - s * 1.8], [a - s * 0.45, b], [a + s * 0.45, b]], "#2B4C39"))];
}
function ecranCouverture() {
  const c = etat.cahier;
  const couv = c.couverture || {};
  const coul = couv.couleur || "#1F4E79";
  const W = 26, H = 35, cx = (1.6 + W - 2.0) / 2;
  const ouvrir = () => {  // au démarrage : l'agenda ; au retour dans l'app : là où l'on était
    const suite = apresCouverture && apresCouverture !== "#/couverture" ? apresCouverture : "#/";
    apresCouverture = null;
    aller(suite);
  };
  // Titre doré : la plus grande taille qui tient (comme sur le PC)
  const titre = c.titre || "Sans titre";
  let taille = 0.6;
  for (const f of [1.5, 1.3, 1.15, 1.0, 0.85, 0.7, 0.6]) {
    if (largeurTexte(titre, `italic ${f * 100}px Georgia, "Times New Roman", serif`) / 100 <= W - 4.8) { taille = f; break; }
  }
  // Image encadrée (ou paysage dessiné tant qu'aucune image n'est choisie)
  const img = couv.image && couv.image.fichier ? couv.image : null;
  const haut = 7.8, bas = H - 5.4;
  let w = 15, hImg = 11;
  if (img) {
    const ratio = Number(img.ratio) || 0.75;
    w = 15.5; hImg = w * ratio;
    if (hImg > bas - haut) { hImg = bas - haut; w = hImg / ratio; }
  }
  const x0 = cx - w / 2, y0 = haut + (bas - haut - hImg) / 2, m = 0.28;
  const zoneImage = svg("g", {});
  if (img) {
    zoneImage.append(svg("rect", {x: x0, y: y0, width: w, height: hImg, fill: "#E9E1CC"}));
    const legende = svg("text", {x: cx, y: y0 + hImg / 2, "text-anchor": "middle", "dominant-baseline": "central",
      "font-size": 0.55, fill: "#8A8270", "font-family": "-apple-system, 'Segoe UI', sans-serif"}, "Chargement de la photo…");
    zoneImage.append(legende);
    urlImage(img.fichier).then((url) => {
      if (url) {
        legende.remove();
        zoneImage.append(svg("image", {href: url, x: x0, y: y0, width: w, height: hImg, preserveAspectRatio: "xMidYMid slice"}));
      } else legende.textContent = etat.mode === "demo" ? "Photo de couverture" : "Photo visible avec le réseau";
    });
  } else zoneImage.append(...illustration(x0, y0, w, hImg).flat(Infinity));
  // Onglets des séparateurs sur la tranche : un toucher ouvre le séparateur
  const secs = c.sections || [];
  const hOnglet = secs.length ? Math.min(4.2, (H - 2.6) / secs.length) : 4;
  const onglets = secs.map((s, k) => {
    const y = 1.6 + k * hOnglet, hh = hOnglet - 0.18;
    const n = Math.max(3, Math.floor(hh / 0.36));
    const nom = s.titre.length > n ? s.titre.slice(0, n - 1) + "…" : s.titre;
    const g = svg("g", {class: "onglet-couverture", role: "button", "aria-label": `Ouvrir ${s.titre}`},
      svg("path", {d: `M ${W - 0.4} ${y} h 1.55 q 0.35 0 0.35 0.35 v ${hh - 0.7} q 0 0.35 -0.35 0.35 h -1.55 z`,
        fill: s.couleur || "#6C7A89", stroke: fonce(s.couleur || "#6C7A89", 0.3), "stroke-width": 0.05}),
      svg("text", {x: W + 0.75, y: y + hh / 2, transform: `rotate(90 ${W + 0.75} ${y + hh / 2})`, "text-anchor": "middle",
        "dominant-baseline": "central", "font-size": 0.56, fill: "#FFFFFF", "font-family": "-apple-system, 'Segoe UI', sans-serif"}, nom));
    g.addEventListener("click", (e) => {
      e.stopPropagation();
      const notes = notesSection(s).length;
      aller(s.type === "calendrier" && !notes ? `#/c/${encodeURIComponent(s.id)}` : `#/s/${encodeURIComponent(s.id)}`);
    });
    return g;
  });
  const anneaux = [];
  for (let y = 1.0; y < H - 0.6; y += 1.36) {
    anneaux.push(svg("path", {d: `M -0.95 ${y + 0.28} C -0.95 ${y - 0.62}, 1.45 ${y - 0.62}, 1.45 ${y + 0.05}`, fill: "none",
      stroke: "#4E545C", "stroke-width": 0.42, "stroke-linecap": "round"}),
    svg("path", {d: `M -0.95 ${y + 0.28} C -0.95 ${y - 0.62}, 1.45 ${y - 0.62}, 1.45 ${y + 0.05}`, fill: "none",
      stroke: "url(#argent)", "stroke-width": 0.24, "stroke-linecap": "round"}),
    svg("rect", {x: 1.18, y: y - 0.1, width: 0.42, height: 0.3, rx: 0.12, fill: "#2B2E33"}));
  }
  const xe = W - 1.8, ew = 0.8;
  const dessin = svg("svg", {viewBox: "-1.4 -0.8 30 37.6", class: "dessin-couverture", role: "img",
    "aria-label": `Couverture de l'agenda « ${titre} » : touchez pour ouvrir`},
    svg("defs", {},
      svg("linearGradient", {id: "argent", x1: 0, y1: 0, x2: 0, y2: 1},
        svg("stop", {offset: "0", "stop-color": "#9AA1A9"}), svg("stop", {offset: "0.45", "stop-color": "#EEF1F3"}),
        svg("stop", {offset: "1", "stop-color": "#8B929A"})),
      svg("linearGradient", {id: "ciel", x1: 0, y1: 0, x2: 0, y2: 1},
        svg("stop", {offset: "0", "stop-color": "#86BFE6"}), svg("stop", {offset: "1", "stop-color": "#FCE3BC"})),
      svg("linearGradient", {id: "elastique", x1: 0, y1: 0, x2: 1, y2: 0},
        svg("stop", {offset: "0", "stop-color": "#141416"}), svg("stop", {offset: "0.5", "stop-color": "#3D3D45"}),
        svg("stop", {offset: "1", "stop-color": "#141416"})),
      svg("clipPath", {id: "forme-couverture"}, svg("rect", {x: 0, y: 0, width: W, height: H, rx: 0.55})),
      svg("filter", {id: "grain", x: 0, y: 0, width: W, height: H, filterUnits: "userSpaceOnUse"},
        svg("feTurbulence", {type: "fractalNoise", baseFrequency: 1.6, numOctaves: 2, seed: 7}),
        svg("feColorMatrix", {type: "matrix", values: "0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 -1.6 1.05"}))),
    // pages sous la couverture, signet
    svg("rect", {x: 0.45, y: 0.35, width: W, height: H, rx: 0.5, fill: "#EFE8D6", stroke: "#CFC4AA", "stroke-width": 0.05}),
    svg("polygon", {points: `14.6,${H - 0.3} 15.4,${H - 0.3} 15.4,${H + 1.5} 15,${H + 1.15} 14.6,${H + 1.5}`, fill: "#A3242B"}),
    onglets,
    // cuir
    svg("rect", {x: 0, y: 0, width: W, height: H, rx: 0.55, fill: coul}),
    svg("rect", {x: 0, y: 0, width: W, height: H, filter: "url(#grain)", "clip-path": "url(#forme-couverture)", opacity: 0.2}),
    svg("rect", {x: 0.03, y: 0.03, width: W - 0.06, height: H - 0.06, rx: 0.55, fill: "none", stroke: fonce(coul, 0.45), "stroke-width": 0.06}),
    svg("rect", {x: 1.6, y: 0.6, width: W - 2.2, height: H - 1.2, rx: 0.55, fill: "none", stroke: melanger(coul, "#F2E6C8", 0.72),
      "stroke-width": 0.07, "stroke-dasharray": "0.32 0.22"}),
    // titre doré et filet orné
    texteDore(cx, 4.6, titre, {"font-size": taille, "font-style": "italic", "font-family": "Georgia, 'Times New Roman', serif"}, coul),
    [-1, 1].map((sens) => svg("line", {x1: cx + sens * 0.5, x2: cx + sens * 4.5, y1: 6.3, y2: 6.3, stroke: COUL_OR, "stroke-width": 0.05})),
    svg("polygon", {points: `${cx},${6.08} ${cx + 0.22},6.3 ${cx},${6.52} ${cx - 0.22},6.3`, fill: COUL_OR, stroke: fonce(coul, 0.3), "stroke-width": 0.03}),
    // image encadrée
    svg("rect", {x: x0 - m + 0.2, y: y0 - m + 0.28, width: w + 2 * m, height: hImg + 2 * m, fill: fonce(coul, 0.5)}),
    svg("rect", {x: x0 - m, y: y0 - m, width: w + 2 * m, height: hImg + 2 * m, fill: "#F4EEDC", stroke: COUL_OR, "stroke-width": 0.06}),
    zoneImage,
    svg("rect", {x: x0, y: y0, width: w, height: hImg, fill: "none", stroke: melanger(COUL_OR, "#000000", 0.25), "stroke-width": 0.04}),
    // année dorée
    texteDore(cx, H - 3.3, String(c.annee || new Date().getFullYear()).split("").join(" "),
      {"font-size": 0.75, "font-family": "Georgia, 'Times New Roman', serif"}, coul),
    // élastique et spirale
    svg("rect", {x: xe + ew, y: -0.34, width: 0.22, height: H + 0.74, fill: fonce(coul, 0.45)}),
    svg("rect", {x: xe, y: -0.34, width: ew, height: H + 0.74, fill: "url(#elastique)"}),
    svg("line", {x1: xe + ew * 0.42, x2: xe + ew * 0.42, y1: -0.34, y2: H + 0.4, stroke: "#55555E", "stroke-width": 0.03}),
    anneaux);
  dessin.addEventListener("click", ouvrir);
  return [h("header", {class: "barre barre-couverture"},
    h("button", {class: "icone", "aria-label": "Réglages", onclick: () => aller("#/reglages")}, "⚙"),
    h("button", {class: "statut", hidden: true}, h("span", {class: "point"}), h("span", {class: "texte"}))),
  h("main", {class: "ecran-couverture"}, dessin,
    h("button", {class: "indice-couverture", onclick: ouvrir}, "Touchez la couverture pour ouvrir l'agenda"))];
}

function ecranAccueil() {
  const c = etat.cahier;
  const auteur = c.modifie_par || {};
  return [barre({titre: c.titre || "Agenda", actions: h("button", {class: "icone", "aria-label": "Réglages", onclick: () => aller("#/reglages")}, "⚙")}),
    ...bandeaux(),
    h("main", {class: "contenu"},
      h("div", {class: "couverture", role: "button", tabindex: "0", "aria-label": "Afficher la couverture",
        style: {background: (c.couverture && c.couverture.couleur) || null}, onclick: () => aller("#/couverture")},
        h("h2", {}, c.titre || "Agenda"),
        h("p", {}, etat.mode === "demo" ? "Démonstration — rien n'est envoyé dans OneDrive"
          : auteur.poste ? `Modifié sur ${auteur.poste === POSTE ? "l'iPhone" : "le poste " + auteur.poste} ${quand(auteur.date)}` : "")),
      h("ul", {class: "liste"}, c.sections.map((s) => {
        const n = notesSection(s).length;
        const cal = s.type === "calendrier";
        return h("li", {}, h("button", {class: "ligne", style: {"--couleur": s.couleur},
          onclick: () => aller(cal && !n ? `#/c/${encodeURIComponent(s.id)}` : `#/s/${encodeURIComponent(s.id)}`)},
          h("span", {class: "pastille"}),
          h("span", {class: "libelle"}, h("b", {}, s.titre), h("small", {}, cal && !n ? "Calendrier de l'année" : `${n} note${n > 1 ? "s" : ""}`)),
          h("span", {class: "chevron"}, "›")));
      })))];
}

function ecranSection(sid) {
  const s = trouverSection(sid);
  if (!s) { setTimeout(() => aller("#/"), 0); return [barre({titre: "Agenda"})]; }
  const entrees = entreesSection(s);
  const modif = peutModifier();
  const lignes = entrees.map((e) => {
    if (e.genre === "calendrier") {
      return h("li", {}, h("button", {class: "ligne", onclick: () => aller(`#/c/${encodeURIComponent(s.id)}`)},
        h("span", {class: "libelle"}, h("b", {}, "Calendrier"), h("small", {}, "Calendrier de l'année et jours fériés")), h("span", {class: "chevron"}, "›")));
    }
    if (e.genre === "sujet") {
      return h("li", {}, h("div", {class: "ligne sujet"},
        h("span", {class: "libelle"}, h("b", {}, e.titre || "Sujet sans titre"), h("small", {}, `${e.notes} note${e.notes > 1 ? "s" : ""}`)),
        modif ? h("button", {class: "plus", "aria-label": "Actions du sujet", onclick: () => menuSujet(s, e.debut)}, "⋯") : null));
    }
    const p = s.pages[e.debut];
    return h("li", {}, h("div", {class: "ligne" + (e.dansSujet ? " dans-sujet" : ""), role: "button", tabindex: "0",
      onclick: (ev) => { if (!ev.target.closest(".plus")) aller(lienNote(s, e.debut, 0)); }},
      h("span", {class: "libelle"}, h("b", {}, e.titre || "Sans titre"),
        h("small", {}, `${libelleType(p)} · ${e.pages} page${e.pages > 1 ? "s" : ""}`)),
      modif ? h("button", {class: "plus", "aria-label": "Actions de la note", onclick: () => menuNoteListe(s, e.debut)}, "⋯") : h("span", {class: "chevron"}, "›")));
  });
  return [barre({titre: s.titre, couleur: s.couleur, retour: "#/"}), ...bandeaux(),
    h("main", {class: "contenu", style: {"--couleur": s.couleur}},
      h("h2", {class: "titre-groupe"}, "Notes de la section"),
      h("ul", {class: "liste"}, lignes.length ? lignes : h("li", {class: "vide"}, "Aucune note pour l'instant")),
      modif ? h("div", {class: "boutons"},
        h("button", {class: "bouton principal", onclick: () => nouvelleNote(s, null)}, "+ Note"),
        s.sujets ? h("button", {class: "bouton", onclick: () => nouveauSujet(s)}, "+ Sujet") : null) : null)];
}

async function menuSujet(s, debut) {
  const titre = s.pages[debut].sujet.titre || "sans titre";
  const r = await choisir(`Sujet « ${titre} »`, [{valeur: "note", libelle: "+ Note dans ce sujet…"},
    {valeur: "renommer", libelle: "Renommer le sujet…"}, {valeur: "supprimer", libelle: "Supprimer le sujet et ses notes…", danger: true}]);
  if (r === "note") nouvelleNote(s, debut);
  else if (r === "renommer") {
    const t = await demanderTexte("Renommer le sujet", "Titre du sujet", s.pages[debut].sujet.titre || "");
    if (t && t.trim()) { s.pages[debut].sujet.titre = t.trim(); modifier(); rendre(); }
  } else if (r === "supprimer") {
    const fin = finSujet(s, debut);
    const notes = notesSection(s).filter(([d]) => debut <= d && d < fin).length;
    if (await confirmer(`Supprimer le sujet « ${titre} », ses ${notes} note${notes > 1 ? "s" : ""} et leurs ${fin - debut} page${fin - debut > 1 ? "s" : ""} ?\n\nCette suppression ne peut pas être annulée.`)) {
      effacerPages(s, debut, fin);
      rendre();
    }
  }
}
async function menuNoteListe(s, debut) {
  const titre = s.pages[debut].note.titre || "sans titre";
  const r = await choisir(`Note « ${titre} »`, [{valeur: "ouvrir", libelle: "Ouvrir la note"},
    {valeur: "renommer", libelle: "Renommer la note…"}, {valeur: "supprimer", libelle: "Supprimer la note…", danger: true}]);
  if (r === "ouvrir") aller(lienNote(s, debut, 0));
  else if (r === "renommer") renommerNote(s, debut);
  else if (r === "supprimer") supprimerNote(s, debut, () => rendre());
}
async function renommerNote(s, debut) {
  const note = s.pages[debut].note;
  const t = await demanderTexte("Renommer la note", "Titre de la note", note.titre || "");
  if (t && t.trim()) { note.titre = t.trim(); modifier(); rendre(); }
}
async function supprimerNote(s, debut, ensuite) {
  const fin = finNote(s, debut), nb = fin - debut;
  const titre = s.pages[debut].note.titre || "sans titre";
  if (!s.pages.slice(debut, fin).every(pageVide)
    && !(await confirmer(`Supprimer la note « ${titre} » (${nb} page${nb > 1 ? "s" : ""}) et tout son contenu ?`))) return;
  effacerPages(s, debut, fin);
  ensuite();
}
/** Supprime des pages et, dans OneDrive, leurs pièces jointes (comme le PC). */
function effacerPages(s, a, b) {
  const supprimees = supprimerPages(s, a, b);
  modifier();
  for (const p of supprimees) if ((p.pieces || []).length) supprimerDansOneDrive("documents", p.id);
}
async function nouvelleNote(s, sujetDebut) {
  if (!peutModifier()) return;
  const sujets = s.sujets ? entreesSection(s).filter((e) => e.genre === "sujet") : [];
  const champs = [{cle: "titre", libelle: "Titre de la note", valeur: ""}];
  if (sujets.length && sujetDebut === null) {
    champs.push({cle: "sujet", libelle: "Sujet", type: "select", valeur: String(sujets[0].debut),
      options: sujets.map((e) => [String(e.debut), e.titre || "Sujet sans titre"])});
  }
  const r = await formulaire("Nouvelle note", champs, "Ajouter");
  if (!r || !r.titre.trim()) return;
  const cible = r.sujet != null ? Number(r.sujet) : sujetDebut;
  const pos = cible !== null ? finSujet(s, cible) : s.pages.length;
  s.pages.splice(pos, 0, nouvellePage(s.modele_page, r.titre.trim()));
  modifier();
  rendre();
  avis(`Note « ${r.titre.trim()} » ajoutée` + (cible !== null ? ` au sujet « ${s.pages[cible].sujet.titre} »` : ""));
}
async function nouveauSujet(s) {
  const r = await formulaire("Nouveau sujet", [{cle: "sujet", libelle: "Titre du sujet", valeur: ""},
    {cle: "note", libelle: "Titre de sa première note", valeur: ""}], "Ajouter");
  if (!r || !r.sujet.trim()) return;
  const p = nouvellePage(s.modele_page, r.note.trim());
  p.sujet = {id: nouvelId(), titre: r.sujet.trim()};
  s.pages.push(p);
  modifier();
  rendre();
}

/* ----- une note et ses pages ----- */
function ecranNote(sid, pid, k) {
  const s = trouverSection(sid);
  const idx = s ? s.pages.findIndex((p) => p.id === pid) : -1;
  const debut = idx >= 0 ? debutNote(s, idx) : null;
  if (debut === null) { setTimeout(() => aller(s ? `#/s/${encodeURIComponent(s.id)}` : "#/"), 0); return [barre({titre: "Agenda"})]; }
  if (debut !== idx) { setTimeout(() => history.replaceState(null, "", lienNote(s, debut, idx - debut + k)), 0); k = idx - debut + k; }
  const fin = finNote(s, debut), n = fin - debut;
  k = Math.max(0, Math.min(k, n - 1));
  const p = s.pages[debut + k];
  const note = s.pages[debut].note;
  const sujet = sujetDe(s, debut);
  const modif = peutModifier();
  const vers = (j) => aller(lienNote(s, debut, j));
  const corps = h("div", {class: "papier", style: {"--couleur": s.couleur}}, editeur(p, s));
  glisser(corps, () => { if (k > 0) vers(k - 1); }, () => { if (k < n - 1) vers(k + 1); });
  const nbPieces = (p.pieces || []).length;
  return [barre({titre: s.titre, couleur: s.couleur, retour: `#/s/${encodeURIComponent(s.id)}`,
    actions: h("button", {class: "icone", "aria-label": "Actions de la page", onclick: () => menuPage(s, debut, k)}, "⋯")}),
  ...bandeaux(),
  h("main", {class: "contenu", style: {"--couleur": s.couleur}},
    h("button", {class: "titre-note", onclick: () => { if (modif) renommerNote(s, debut); }},
      sujet !== null ? h("span", {class: "sujet-de"}, s.pages[sujet].sujet.titre || "Sujet") : null,
      note.titre || "Sans titre"),
    h("div", {class: "barre-page"},
      h("button", {class: "nav", "aria-label": "Page précédente", disabled: k === 0, onclick: () => vers(k - 1)}, "‹"),
      h("span", {class: "numero"}, `Page ${k + 1} sur ${n} · ${libelleType(p)}`),
      h("button", {class: "nav", "aria-label": "Page suivante", disabled: k >= n - 1, onclick: () => vers(k + 1)}, "›")),
    corps,
    h("div", {class: "pied-page"},
      h("span", {}, p.cree ? "Créée le " + dateCourte(p.cree) : ""),
      h("span", {}, "Pièces jointes : ", h("button", {class: "lien", onclick: () => ouvrirPieces(p)}, String(nbPieces)),
        " Vocal : ", h("button", {class: "lien", onclick: () => ouvrirVocaux(p)}, String((p.vocaux || []).length)))))];
}

function glisser(el, precedente, suivante) {
  let x0 = 0, y0 = 0, t0 = 0, actif = false;
  el.addEventListener("touchstart", (e) => {
    const c = e.target;
    actif = e.touches.length === 1 && !c.closest("input, select, button, a, .grille-defil")
      && !(c.tagName === "TEXTAREA" && document.activeElement === c);
    x0 = e.touches[0].clientX; y0 = e.touches[0].clientY; t0 = Date.now();
  }, {passive: true});
  el.addEventListener("touchend", (e) => {
    if (!actif) return;
    const t = e.changedTouches[0], dx = t.clientX - x0, dy = t.clientY - y0;
    if (Math.abs(dx) > 70 && Math.abs(dy) < 50 && Date.now() - t0 < 700) (dx > 0 ? precedente : suivante)();
  }, {passive: true});
}

async function menuPage(s, debut, k) {
  const fin = finNote(s, debut), n = fin - debut, pi = debut + k, p = s.pages[pi];
  const modif = peutModifier();
  const r = await choisir(`Page ${k + 1} de « ${s.pages[debut].note.titre || "sans titre"} »`, [
    modif ? {valeur: "page", libelle: "+ Page…"} : null,
    modif && n > 1 ? {valeur: "deplacer", libelle: "Déplacer la page…"} : null,
    {valeur: "pieces", libelle: `Pièces jointes (${(p.pieces || []).length})…`},
    modif ? {valeur: "renommer", libelle: "Renommer la note…"} : null,
    modif ? {valeur: "suppr-page", libelle: "Supprimer la page", danger: true} : null,
    modif ? {valeur: "suppr-note", libelle: "Supprimer la note…", danger: true} : null]);
  if (r === "page") ajouterPage(s, debut, k);
  else if (r === "deplacer") {
    const c = await choisir("Déplacer la page (elle reste dans sa note)", [
      k > 0 ? {valeur: "avant", libelle: "Déplacer la page avant"} : null,
      k < n - 1 ? {valeur: "apres", libelle: "Déplacer la page après"} : null,
      k > 0 ? {valeur: "debut", libelle: "Déplacer la page au début"} : null,
      k < n - 1 ? {valeur: "fin", libelle: "Déplacer la page à la fin"} : null]);
    if (c) {
      const nk = deplacerPage(s, debut, fin, pi, c);
      if (nk !== null) { modifier(); aller(lienNote(s, debut, nk - debut)); }
    }
  } else if (r === "pieces") ouvrirPieces(p);
  else if (r === "renommer") renommerNote(s, debut);
  else if (r === "suppr-page") {
    if (!pageVide(p) && !(await confirmer(`Supprimer la page ${k + 1} et tout son contenu ?`))) return;
    effacerPages(s, pi, pi + 1);
    const reste = n - 1;
    if (!reste) aller(`#/s/${encodeURIComponent(s.id)}`);
    else aller(lienNote(s, debut, Math.min(k, reste - 1)));
  } else if (r === "suppr-note") supprimerNote(s, debut, () => aller(`#/s/${encodeURIComponent(s.id)}`));
}
async function ajouterPage(s, debut, k) {
  const p = s.pages[debut + k];
  const actuel = p.type === "web" ? "web:" + p.service : p.type;
  const r = await formulaire("Nouvelle page", [
    {cle: "type", libelle: "Type de page", type: "select", valeur: TYPES_LIBELLES.some(([v]) => v === actuel) ? actuel : "lignee",
      options: TYPES_LIBELLES},
    {cle: "position", libelle: "Position", type: "select", valeur: "suite",
      options: [["suite", "Ajouter à la suite"], ["debut", "Ajouter au début"], ["fin", "Ajouter en dernier"]]}], "Ajouter");
  if (!r) return;
  let modele = {type: r.type};
  if (r.type.startsWith("web:")) modele = {type: "web", service: r.type.slice(4)};
  else if (r.type === "colonnes") {
    const t = await demanderTexte("Colonnes", "Titres des colonnes, séparés par des virgules",
      (p.colonnes || ["Colonne 1", "Colonne 2"]).join(", "));
    const colonnes = (t || "").split(",").map((x) => x.trim()).filter(Boolean).slice(0, 8);
    if (!colonnes.length) return;
    modele.colonnes = colonnes;
    if (p.type === "colonnes" && p.largeurs && p.largeurs.length === colonnes.length) modele.largeurs = p.largeurs.slice();
  } else if (r.type === "tableur") {
    const url = await demanderTexte("Tableur", "Adresse de la feuille Google Sheets", URL_TABLEUR);
    if (!url) return;
    modele.url = url.includes("://") ? url : "https://" + url;
  }
  const fin = finNote(s, debut);
  const nouvelle = ajouterPageNote(s, debut, fin, debut + k, modele, r.position);
  modifier();
  const pos = s.pages.indexOf(nouvelle);
  aller(lienNote(s, debut, pos - debut));
}

/* ----- éditeurs selon le type de page ----- */
function editeur(p, s) {
  switch (p.type) {
    case "lignee": case "quadrillee": return editeurTexte(p);
    case "blanche": return editeurDessin(p);
    case "colonnes": return editeurColonnes(p);
    case "finance": return editeurFinance(p);
    case "todo": return editeurTodo(p);
    case "agenda": return editeurAgenda(p, s);
    case "web": case "tableur": return editeurWeb(p);
    default: return h("p", {class: "note-info"}, "Ce type de page se consulte sur le PC.");
  }
}
function ajusterHauteur(ta) { ta.style.height = "auto"; ta.style.height = ta.scrollHeight + 4 + "px"; }

function editeurTexte(p) {
  p.texte = p.texte || "";
  const ta = h("textarea", {class: "zone-texte" + (p.type === "quadrillee" ? " quadrillee" : ""), value: p.texte,
    disabled: !peutModifier(), autocapitalize: "sentences", spellcheck: "true", "aria-label": "Texte de la page",
    style: stylePolice(p, 17)});
  let precedent = ta.value;
  ta.addEventListener("input", () => {
    const texte = appliquerSaisieTexte(p, precedent, ta.value);
    if (texte !== ta.value) {
      const pos = Math.max(0, ta.selectionStart - (ta.value.length - texte.length));
      ta.value = texte;
      ta.setSelectionRange(pos, pos);
    }
    precedent = texte;
    ajusterHauteur(ta);
    modifier();
  });
  const images = p.images || [];
  return [ta, images.length ? h("p", {class: "note-info"}, `Les ${OBJ} du texte marquent la place des images de la page :`) : null,
    images.length ? h("div", {class: "images-page"}, images.map((im) => figureImage(im.fichier))) : null];
}
function figureImage(fichier) {
  const img = h("img", {alt: "", loading: "lazy"});
  const legende = h("figcaption", {}, "Chargement…");
  urlImage(fichier).then((url) => {
    if (url) { img.src = url; legende.remove(); } else legende.textContent = etat.mode === "demo" ? "Image (démonstration)" : "Image visible avec le réseau";
  });
  return h("figure", {}, img, legende);
}
function editeurDessin(p) {
  const zone = svg("svg", {viewBox: "0 0 26 35", class: "dessin", role: "img", "aria-label": "Dessin de la page"});
  for (const im of p.images_libres || []) {
    const el = svg("image", {x: im.x, y: im.y, width: im.largeur, height: im.largeur * (im.ratio || 0.75), opacity: im.opacite, preserveAspectRatio: "none"});
    zone.append(el);
    urlImage(im.fichier).then((url) => { if (url) el.setAttribute("href", url); });
  }
  for (const tr of p.traits || []) {
    const pts = tr.p || [];
    let points = [];
    for (let i = 0; i + 1 < pts.length; i += 2) points.push(`${pts[i]},${pts[i + 1]}`);
    if (points.length === 1) points.push(`${pts[0] + 0.001},${pts[1]}`);
    zone.append(svg("polyline", {points: points.join(" "), fill: "none", stroke: tr.c || "#1D2B53", "stroke-width": tr.e || 0.09,
      "stroke-linecap": "round", "stroke-linejoin": "round"}));
  }
  return [zone, h("p", {class: "note-info"}, "Page de dessin : elle se consulte ici et se modifie sur le PC.")];
}
function editeurColonnes(p) {
  p.colonnes = p.colonnes && p.colonnes.length ? p.colonnes : ["Colonne 1"];
  p.cellules = p.cellules || [];
  const modif = peutModifier();
  const tel = p.colonnes.map((c) => /^t[ée]l/i.test(c.trim()));
  const corps = h("tbody");
  const police = stylePolice(p, 17);
  const valeur = (i, k) => (i < p.cellules.length && k < p.cellules[i].length ? p.cellules[i][k] : "");
  const remplir = () => {
    const nb = Math.min(NB_CASES, p.cellules.length + (modif ? 1 : 0));
    corps.replaceChildren();
    for (let i = 0; i < nb; i++) corps.append(rangee(i));
    info.textContent = p.cellules.length >= NB_CASES ? `Page pleine (${NB_CASES} lignes) : ajoutez une page (⋯ › + Page) pour continuer.` : "";
  };
  const rangee = (i) => {
    const numeros = p.colonnes.map((_, k) => valeur(i, k)).filter((v, k) => tel[k] && v.trim());
    return h("tr", {}, p.colonnes.map((_, k) => h("td", {}, h("input", {value: valeur(i, k), disabled: !modif, style: police,
      inputmode: tel[k] ? "tel" : null, autocapitalize: tel[k] ? "off" : "words", "aria-label": p.colonnes[k],
      oninput: (e) => {
        const avant = p.cellules.length;
        poserCellule(p, i, k, e.target.value);
        modifier();
        if (p.cellules.length > avant && i === corps.children.length - 1 && corps.children.length < NB_CASES) corps.append(rangee(i + 1));
      }}))),
    h("td", {class: "action"}, numeros.length ? h("a", {class: "appel", href: "tel:" + numeros[0].replace(/[^\d+]/g, ""), "aria-label": "Appeler"}, "📞")
      : modif && i < p.cellules.length ? h("button", {class: "lien", "aria-label": "Effacer la ligne", onclick: async () => {
        if (await confirmer("Effacer cette ligne ?", "Effacer")) { p.cellules.splice(i, 1); poserCellule(p, 0, 0, valeur(0, 0)); modifier(); remplir(); }
      }}, "✕") : null));
  };
  const info = h("p", {class: "note-info"});
  remplir();
  const largeurs = largeursColonnes(p.colonnes.length, p.largeurs);  // réglées sur le PC (Modifier les colonnes)
  return [h("div", {class: "grille-defil"}, h("table", {class: "grille" + (largeurs ? " proportionnelle" : "")},
    largeurs ? h("colgroup", {}, largeurs.map((v) => h("col", {style: {width: `${(v * 0.88).toFixed(2)}%`}})),
      h("col", {style: {width: "12%"}})) : null,  // 12 % : le bouton d'appel
    h("thead", {}, h("tr", {}, p.colonnes.map((c) => h("th", {}, c)), h("th", {}))), corps)), info];
}

function editeurFinance(p) {
  p.compte = p.compte || {solde_initial: 0};
  p.operations = p.operations || [];
  p.modeles = p.modeles || [];
  const modif = peutModifier();
  const {triees, soldes: liste} = soldes(p);
  const initial = Number(p.compte.solde_initial) || 0;
  const final = liste.length ? liste[liste.length - 1] : initial;
  const debits = triees.reduce((t, o) => t + (+o.debit || 0), 0), credits = triees.reduce((t, o) => t + (+o.credit || 0), 0);
  const pointees = triees.filter((o) => o.pointe);
  const debitsP = pointees.reduce((t, o) => t + (+o.debit || 0), 0), creditsP = pointees.reduce((t, o) => t + (+o.credit || 0), 0);
  const pointe = Math.round((initial + creditsP - debitsP) * 100) / 100;
  const ouverture = h("input", {class: "saisie", value: montantTexte(initial), inputmode: "decimal", disabled: !modif, "aria-label": "Solde d'ouverture",
    onchange: (e) => {
      const v = lireMontant(e.target.value);
      if (v === null) { avis("Montant illisible"); e.target.value = montantTexte(initial); return; }
      p.compte.solde_initial = v;
      modifier();
      rendre();
    }});
  return [
    h("div", {class: "solde-carte"}, h("span", {}, "Solde"), h("span", {class: "grand" + (final < 0 ? " negatif" : "")}, montantTexte(final, true))),
    h("div", {class: "ouverture"}, h("label", {}, "Solde d'ouverture"), ouverture),
    h("ul", {class: "operations"}, triees.map((o, i) => h("li", {class: "op"},
      h("input", {type: "checkbox", class: "pointe", checked: !!o.pointe, disabled: !modif, "aria-label": "Pointée",
        onchange: (e) => { o.pointe = e.target.checked; modifier(); rendre(); }}),
      h("button", {class: "corps", onclick: () => { if (modif) editerOperation(p, o); }},
        h("b", {}, o.tiers || "(sans tiers)"), h("small", {}, dateCourte(o.date))),
      h("div", {class: "montants"},
        +o.debit ? h("span", {class: "mvt"}, "− " + montantTexte(+o.debit, true)) : null,
        +o.credit ? h("span", {class: "mvt credit"}, "+ " + montantTexte(+o.credit, true)) : null,
        h("small", {class: liste[i] < 0 ? "negatif" : ""}, montantTexte(liste[i], true)))))),
    triees.length ? null : h("p", {class: "note-info"}, "Aucune opération pour l'instant."),
    modif ? h("div", {class: "boutons"}, h("button", {class: "bouton principal", onclick: () => editerOperation(p, null)}, "+ Opération")) : null,
    h("table", {class: "totaux"},
      h("tr", {}, h("th", {}), h("th", {}, "Débit"), h("th", {}, "Crédit"), h("th", {}, "Solde")),
      h("tr", {}, h("td", {}, "Totaux"), h("td", {}, montantTexte(debits)), h("td", {}, montantTexte(credits)),
        h("td", {class: final < 0 ? "negatif" : ""}, montantTexte(final, true))),
      h("tr", {}, h("td", {}, "Totaux pointés"), h("td", {}, montantTexte(debitsP)), h("td", {}, montantTexte(creditsP)),
        h("td", {class: pointe < 0 ? "negatif" : ""}, montantTexte(pointe, true)))),
    h("details", {class: "modeles"}, h("summary", {}, `Lignes préenregistrées (${p.modeles.length})`),
      h("ul", {class: "liste"}, p.modeles.length ? p.modeles.map((m) => h("li", {}, h("button", {class: "ligne", onclick: () => { if (modif) editerModele(p, m); }},
        h("span", {class: "libelle"}, h("b", {}, m.libelle), h("small", {}, `${m.sens === "credit" ? "Crédit" : "Débit"}${m.montant ? " · " + montantTexte(m.montant, true) : ""}`)),
        h("span", {class: "chevron"}, "›")))) : h("li", {class: "vide"}, "Aucune ligne préenregistrée")),
      modif ? h("div", {class: "boutons"}, h("button", {class: "bouton", onclick: () => editerModele(p, null)}, "+ Ligne préenregistrée")) : null)];
}
function trouverModele(p, libelle) {
  const cle = sansAccents(libelle.trim());
  return cle ? (p.modeles || []).find((m) => sansAccents(m.libelle) === cle) || null : null;
}
async function editerOperation(p, o) {
  const listeId = "modeles-" + p.id;
  const r = await ouvrirFeuille((fermer) => {
    const date = h("input", {type: "date", value: o ? o.date : iso(aujourdhui())});
    const tiers = h("input", {value: o ? o.tiers : "", list: listeId, autocapitalize: "words"});
    const debit = h("input", {value: o && +o.debit ? montantTexte(+o.debit) : "", inputmode: "decimal", placeholder: "0,00"});
    const credit = h("input", {value: o && +o.credit ? montantTexte(+o.credit) : "", inputmode: "decimal", placeholder: "0,00"});
    const pointee = h("input", {type: "checkbox", checked: !!(o && o.pointe), style: {width: "24px", height: "24px"}});
    const appliquer = (m) => {
      if (!m) return;
      tiers.value = m.libelle;
      if (m.montant) {
        (m.sens === "credit" ? credit : debit).value = montantTexte(m.montant);
        (m.sens === "credit" ? debit : credit).value = "";
      }
    };
    tiers.addEventListener("change", () => { const m = trouverModele(p, tiers.value); if (m && !debit.value && !credit.value) appliquer(m); });
    const enregistrer = () => {
      const d = debit.value.trim() ? lireMontant(debit.value) : 0, c = credit.value.trim() ? lireMontant(credit.value) : 0;
      if (!dateIso(date.value)) { avis("Date invalide"); return; }
      if (d === null || c === null) { avis("Montant illisible"); return; }
      fermer({date: date.value, tiers: tiers.value.trim(), debit: Math.abs(d), credit: Math.abs(c), pointe: pointee.checked});
    };
    return [h("h2", {}, o ? "Opération" : "Nouvelle opération"),
      p.modeles.length ? h("label", {class: "champ"}, h("span", {}, "Ligne préenregistrée"),
        h("select", {onchange: (e) => appliquer(p.modeles.find((m) => m.id === e.target.value))},
          h("option", {value: ""}, "Choisir…"), p.modeles.map((m) => h("option", {value: m.id}, m.libelle)))) : null,
      h("label", {class: "champ"}, h("span", {}, "Date"), date),
      h("label", {class: "champ"}, h("span", {}, "Tiers"), tiers),
      h("datalist", {id: listeId}, p.modeles.map((m) => h("option", {value: m.libelle}))),
      h("div", {class: "deux"}, h("label", {class: "champ"}, h("span", {}, "Débit"), debit), h("label", {class: "champ"}, h("span", {}, "Crédit"), credit)),
      h("label", {class: "champ", style: {display: "flex", gap: "10px", alignItems: "center"}}, pointee, h("span", {style: {margin: 0}}, "Pointée")),
      h("div", {class: "actions"},
        o ? h("button", {class: "bouton danger", onclick: () => fermer("supprimer")}, "Supprimer") : null,
        h("button", {class: "bouton", onclick: () => fermer(null)}, "Annuler"),
        h("button", {class: "bouton principal", onclick: enregistrer}, "Enregistrer"))];
  });
  if (!r) return;
  if (r === "supprimer") {
    if (!(await confirmer(`Supprimer l'opération « ${o.tiers || "sans tiers"} » du ${dateCourte(o.date)} ?`))) return;
    p.operations.splice(p.operations.indexOf(o), 1);
  } else if (o) Object.assign(o, r);
  else p.operations.push(Object.assign(nouvelleOperation(r.date), r));
  modifier();
  rendre();
}
async function editerModele(p, m) {
  const r = await ouvrirFeuille((fermer) => {
    let sens = m ? m.sens : "debit";
    const libelle = h("input", {value: m ? m.libelle : "", autocapitalize: "words"});
    const montant = h("input", {value: m && m.montant ? montantTexte(m.montant) : "", inputmode: "decimal", placeholder: "0,00"});
    const boutons = ["debit", "credit"].map((v) => h("button", {class: v === sens ? "choisi" : "", onclick: () => {
      sens = v;
      boutons.forEach((b, i) => b.classList.toggle("choisi", ["debit", "credit"][i] === v));
    }}, v === "debit" ? "Débit" : "Crédit"));
    return [h("h2", {}, "Ligne préenregistrée"),
      h("label", {class: "champ"}, h("span", {}, "Libellé (tiers)"), libelle),
      h("div", {class: "champ"}, h("span", {}, "Sens"), h("div", {class: "segments"}, boutons)),
      h("label", {class: "champ"}, h("span", {}, "Montant habituel"), montant),
      h("div", {class: "actions"},
        m ? h("button", {class: "bouton danger", onclick: () => fermer("supprimer")}, "Supprimer") : null,
        h("button", {class: "bouton", onclick: () => fermer(null)}, "Annuler"),
        h("button", {class: "bouton principal", onclick: () => {
          const v = montant.value.trim() ? lireMontant(montant.value) : 0;
          if (!libelle.value.trim()) { avis("Le libellé est requis"); return; }
          if (v === null) { avis("Montant illisible"); return; }
          fermer({libelle: libelle.value.trim(), sens, montant: Math.abs(v)});
        }}, "Enregistrer"))];
  });
  if (!r) return;
  if (r === "supprimer") p.modeles.splice(p.modeles.indexOf(m), 1);
  else if (m) Object.assign(m, r);
  else p.modeles.push(Object.assign({id: nouvelId()}, r));
  trierModeles(p);
  modifier();
  rendre();
}

function editeurTodo(p) {
  p.taches = p.taches || [];
  const modif = peutModifier();
  const auj = iso(aujourdhui());
  const liste = h("ul", {class: "taches"});
  const ligne = (t) => {
    const echeance = h("div", {class: "echeance" + (!t.fait && t.echeance && t.echeance < auj ? " retard" : "")}, "Échéance :",
      h("input", {type: "date", value: t.echeance || "", disabled: !modif, "aria-label": "Échéance",
        onchange: (e) => {
          t.echeance = dateIso(e.target.value) ? e.target.value : null;
          echeance.classList.toggle("retard", !t.fait && !!t.echeance && t.echeance < auj);
          modifier();
        }}));
    const li = h("li", {class: "tache" + (t.fait ? " faite" : "")},
      h("input", {type: "checkbox", checked: !!t.fait, disabled: !modif, "aria-label": "Faite",
        onchange: (e) => {
          t.fait = e.target.checked;
          li.classList.toggle("faite", t.fait);
          echeance.classList.toggle("retard", !t.fait && !!t.echeance && t.echeance < auj);
          modifier();
          bilan.textContent = texteBilan();
        }}),
      h("input", {class: "texte", value: t.texte, disabled: !modif, placeholder: "Tâche", "aria-label": "Tâche", autocapitalize: "sentences",
        oninput: (e) => { t.texte = e.target.value; modifier(); }}),
      modif ? h("button", {class: "suppr", "aria-label": "Supprimer la tâche", onclick: () => {
        p.taches.splice(p.taches.indexOf(t), 1);
        li.remove();
        modifier();
        bilan.textContent = texteBilan();
      }}, "✕") : null,
      echeance);
    return li;
  };
  const texteBilan = () => { const f = p.taches.filter((t) => t.fait).length; return `${p.taches.length} tâche${p.taches.length > 1 ? "s" : ""}, ${f} faite${f > 1 ? "s" : ""}`; };
  const bilan = h("p", {class: "note-info"}, texteBilan());
  for (const t of p.taches) liste.append(ligne(t));
  return [liste, bilan, modif ? h("div", {class: "boutons"},
    h("button", {class: "bouton principal", onclick: () => {
      const t = nouvelleTache("");
      p.taches.push(t);
      const li = ligne(t);
      liste.append(li);
      li.querySelector(".texte").focus();
      modifier();
      bilan.textContent = texteBilan();
    }}, "+ Tâche"),
    h("button", {class: "bouton", onclick: async () => {
      const n = p.taches.filter((t) => t.fait).length;
      if (!n) { avis("Aucune tâche faite"); return; }
      if (!(await confirmer(`Retirer les ${n} tâche${n > 1 ? "s" : ""} faite${n > 1 ? "s" : ""} ?`, "Retirer"))) return;
      p.taches = p.taches.filter((t) => !t.fait);
      modifier();
      rendre();
    }}, "Retirer les tâches faites")) : null];
}

function editeurAgenda(p, s) {
  p.jours = p.jours && typeof p.jours === "object" ? p.jours : {};
  const modif = peutModifier();
  const lundi = !!etat.cahier.semaine_lundi;
  const debut = semaines.get(p.id) || debutSemaine(aujourdhui(), lundi);
  const fin = ajouterJours(debut, 6);
  const changer = (d) => { semaines.set(p.id, d === 0 ? debutSemaine(aujourdhui(), lundi) : ajouterJours(debut, 7 * d)); rendre(); };
  const titre = debut.getMonth() === fin.getMonth()
    ? `Du ${debut.getDate()} au ${fin.getDate()} ${MOIS[fin.getMonth()]} ${fin.getFullYear()}`
    : `Du ${debut.getDate()} ${MOIS[debut.getMonth()]} au ${fin.getDate()} ${MOIS[fin.getMonth()]} ${fin.getFullYear()}`;
  const auj = iso(aujourdhui());
  const jours = [];
  for (let j = 0; j < 7; j++) {
    const d = ajouterJours(debut, j), cle = iso(d), nom = ferie(d);
    const lignes = p.jours[cle] || [];
    jours.push(h("div", {class: "jour" + (cle === auj ? " aujourdhui" : "")},
      h("h3", {}, dateLongue(d), nom ? h("span", {class: "ferie"}, nom) : null),
      Array.from({length: LIGNES_JOUR}, (_, k) => h("input", {value: lignes[k] || "", disabled: !modif, autocapitalize: "sentences",
        "aria-label": `${dateLongue(d)}, ligne ${k + 1}`,
        oninput: (e) => {
          const l = (p.jours[cle] || []).slice();
          while (l.length <= k) l.push("");
          l[k] = e.target.value;
          while (l.length && !l[l.length - 1].trim()) l.pop();
          if (l.length) p.jours[cle] = l; else delete p.jours[cle];
          modifier();
        }}))));
  }
  const corps = h("div", {}, h("div", {class: "semaine-nav"},
    h("button", {class: "bouton", "aria-label": "Semaine précédente", onclick: () => changer(-1)}, "‹"),
    h("span", {class: "titre"}, titre),
    h("button", {class: "bouton", "aria-label": "Semaine suivante", onclick: () => changer(1)}, "›")),
  jours,
  h("div", {class: "boutons"}, h("button", {class: "bouton", onclick: () => changer(0)}, "Cette semaine")));
  return corps;
}

function editeurWeb(p) {
  const [nom, defaut] = p.type === "tableur" ? ["Google Sheets", URL_TABLEUR] : SERVICES[p.service] || SERVICES.page;
  const url = /^https?:\/\//i.test(p.url || "") ? p.url : defaut;   // seulement une adresse web
  return h("div", {class: "page-web"},
    h("div", {class: "service"}, nom),
    h("p", {}, "Les pages Google s'ouvrent dans Safari : elles ne peuvent pas s'afficher à l'intérieur de l'app."),
    h("a", {class: "bouton principal", href: url, target: "_blank", rel: "noopener"}, `Ouvrir ${nom}`));
}

/* ----- calendrier annuel ----- */
function ecranCalendrier(sid) {
  const s = trouverSection(sid);
  if (!s) { setTimeout(() => aller("#/"), 0); return [barre({titre: "Agenda"})]; }
  const annee = anneeCalendrier || new Date().getFullYear();
  const lundi = !!etat.cahier.semaine_lundi;
  const auj = iso(aujourdhui());
  const agenda = (() => {
    for (const sec of etat.cahier.sections) for (const p of sec.pages) if (p.type === "agenda") return {sec, p};
    return null;
  })();
  const ouvrirJour = (d) => {
    if (!agenda) return;
    semaines.set(agenda.p.id, debutSemaine(d, lundi));
    const pi = agenda.sec.pages.indexOf(agenda.p), debut = debutNote(agenda.sec, pi);
    if (debut !== null) aller(lienNote(agenda.sec, debut, pi - debut));
  };
  const tetes = lundi ? JOURS_COURTS : [JOURS_COURTS[6], ...JOURS_COURTS.slice(0, 6)];
  const mois = [];
  for (let m = 0; m < 12; m++) {
    const premier = new Date(annee, m, 1);
    const decalage = lundi ? jourSemaine(premier) : premier.getDay();
    const cases = tetes.map((t) => h("span", {class: "tete"}, t));
    for (let i = 0; i < decalage; i++) cases.push(h("span", {}));
    for (let j = 1; j <= new Date(annee, m + 1, 0).getDate(); j++) {
      const d = new Date(annee, m, j), nom = ferie(d);
      const cls = [jourSemaine(d) >= 5 ? "fds" : "", nom ? "ferie" : "", iso(d) === auj ? "auj" : ""].join(" ").trim();
      cases.push(h("button", {class: cls, title: nom || null, onclick: () => ouvrirJour(d)}, String(j)));
    }
    mois.push(h("div", {class: "mois"}, h("h3", {}, MOIS[m]), h("div", {class: "jours"}, cases)));
  }
  const feries = etat.cahier.feries === false ? [] : [...feriesQuebec(annee)];
  return [barre({titre: s.titre, couleur: s.couleur, retour: notesSection(s).length ? `#/s/${encodeURIComponent(s.id)}` : "#/"}), ...bandeaux(),
    h("main", {class: "contenu"},
      h("div", {class: "annee-nav"},
        h("button", {class: "bouton", "aria-label": "Année précédente", onclick: () => { anneeCalendrier = annee - 1; rendre(); }}, "‹"),
        h("span", {class: "titre"}, String(annee)),
        h("button", {class: "bouton", "aria-label": "Année suivante", onclick: () => { anneeCalendrier = annee + 1; rendre(); }}, "›")),
      agenda ? h("p", {class: "note-info", style: {marginTop: 0}}, "Touchez un jour pour ouvrir sa semaine dans l'agenda.") : null,
      h("div", {class: "mois-grille"}, mois),
      feries.length ? [h("h2", {class: "titre-groupe"}, `Jours fériés ${annee}`),
        h("ul", {class: "liste liste-feries"}, feries.map(([d, nom]) => h("li", {}, h("span", {}, dateLongue(dateIso(d))), h("span", {}, nom))))] : null)];
}

/* ----- pièces jointes ----- */
/* ----- messages vocaux (enregistrés sur le PC avec « + Vocal ») ----- */
function texteDuree(secondes) {
  const total = Math.round(Number(secondes) || 0);
  return `${Math.floor(total / 60)}:${String(total % 60).padStart(2, "0")}`;
}
async function ouvrirVocaux(p) {
  const enLigne = etat.mode === "reel" && !!etat.dossier && !!jetons && navigator.onLine;
  const modif = peutModifier();
  const r = await ouvrirFeuille((fermer) => {
    const vocaux = p.vocaux || [];
    const liste = h("ul", {class: "liste pieces"}, vocaux.length ? vocaux.map((x) => {
      const audio = h("audio", {controls: true, preload: "none", class: "lecteur-vocal"});
      const note = h("small", {class: "note-info"}, etat.mode === "demo" ? "Démonstration : pas de fichier à écouter"
        : enLigne ? "" : "L'écoute demande le réseau");
      if (enLigne) {
        adresseTelechargement(["vocal", p.id, x.fichier]).then((url) => { if (url) audio.src = url; else note.textContent = "Fichier introuvable"; })
          .catch(() => { note.textContent = "Indisponible"; });
      }
      return h("li", {}, h("div", {class: "ligne vocal"},
        h("span", {class: "libelle"}, h("b", {}, x.description || x.fichier),
          h("small", {}, `${texteDuree(x.duree)}${x.date ? " · " + dateCourte(String(x.date).slice(0, 10)) : ""}`)),
        modif ? h("button", {class: "plus", "aria-label": "Supprimer le message", onclick: () => fermer({supprimer: x})}, "✕") : null,
        enLigne ? audio : null, note));
    }) : h("li", {class: "vide"}, "Aucun message vocal"));
    return [h("h2", {}, "Messages vocaux"), liste,
      h("p", {class: "note-info"}, "Les messages s'enregistrent sur le PC avec « + Vocal »."),
      h("div", {class: "actions"}, h("button", {class: "bouton", onclick: () => fermer(null)}, "Fermer"))];
  });
  if (r && r.supprimer && await confirmer(`Supprimer le message vocal « ${r.supprimer.description || r.supprimer.fichier} » ?`)) {
    p.vocaux.splice(p.vocaux.indexOf(r.supprimer), 1);
    modifier();
    supprimerDansOneDrive("vocal", p.id, r.supprimer.fichier);
    rendre();
    return ouvrirVocaux(p);
  }
}
async function ouvrirPieces(p) {
  const enLigne = etat.mode === "reel" && !!etat.dossier && !!jetons && navigator.onLine;
  const modif = peutModifier();
  const adresses = new Map();
  const r = await ouvrirFeuille((fermer) => {
    const pieces = p.pieces || [];
    const liste = h("ul", {class: "liste pieces"}, pieces.length ? pieces.map((x) => {
      const lien = h("a", {class: "bouton", href: "#", target: "_blank", rel: "noopener", "aria-disabled": "true",
        onclick: (e) => { if (!adresses.get(x.id)) { e.preventDefault(); avis(etat.mode === "demo" ? "Démonstration : pas de fichier à ouvrir" : enLigne ? "Lien en préparation…" : "Hors ligne : le fichier s'ouvre avec le réseau"); } }}, "Ouvrir");
      if (enLigne) {
        adresseTelechargement(["documents", p.id, x.fichier]).then((url) => {
          if (url) { adresses.set(x.id, url); lien.href = url; lien.removeAttribute("aria-disabled"); } else lien.textContent = "Introuvable";
        }).catch(() => { lien.textContent = "Indisponible"; });
      }
      return h("li", {}, h("div", {class: "ligne"},
        h("span", {class: "libelle"}, h("b", {}, x.description || x.fichier), h("small", {}, x.fichier)),
        lien,
        modif ? h("button", {class: "plus", "aria-label": "Supprimer la pièce jointe", onclick: () => fermer({supprimer: x})}, "✕") : null));
    }) : h("li", {class: "vide"}, "Aucune pièce jointe"));
    const choix = h("input", {type: "file", hidden: true, onchange: (e) => { if (e.target.files[0]) fermer({ajouter: e.target.files[0]}); }});
    return [h("h2", {}, "Pièces jointes"), liste, choix, h("div", {class: "actions"},
      h("button", {class: "bouton", onclick: () => fermer(null)}, "Fermer"),
      modif ? h("button", {class: "bouton principal", onclick: () => {
        if (etat.mode === "reel" && !enLigne) { avis("Hors ligne : l'ajout d'un fichier demande le réseau"); return; }
        choix.click();
      }}, "Ajouter…") : null)];
  });
  if (!r) return;
  if (r.supprimer) {
    if (await confirmer(`Supprimer la pièce jointe « ${r.supprimer.description || r.supprimer.fichier} » ?`)) {
      p.pieces.splice(p.pieces.indexOf(r.supprimer), 1);
      modifier();
      supprimerDansOneDrive("documents", p.id, r.supprimer.fichier);
      rendre();
    }
    return ouvrirPieces(p);
  }
  if (r.ajouter) {
    const f = r.ajouter;
    const point = f.name.lastIndexOf(".");
    const description = await demanderTexte("Pièce jointe", "Description", point > 0 ? f.name.slice(0, point) : f.name);
    if (description === null) return ouvrirPieces(p);
    let nom = f.name;
    if (etat.mode === "reel") {
      avis("Envoi du fichier dans OneDrive…", 60000);
      try { nom = await envoyerPiece(p.id, f); } catch (e) {
        avis("");
        await message("Pièce jointe", "Le fichier n'a pas pu être envoyé dans OneDrive.\n\n" + (e.message || e));
        return ouvrirPieces(p);
      }
      avis("Fichier envoyé");
    }
    (p.pieces || (p.pieces = [])).push({id: nouvelId(), description: description.trim(), fichier: nom});
    modifier();
    rendre();
    return ouvrirPieces(p);
  }
}

/* ----- réglages ----- */
function ecranReglages() {
  const ligne = (libelle, valeur) => h("li", {class: "ligne"}, h("span", {class: "libelle"}, h("b", {}, libelle)), h("span", {class: "valeur"}, valeur));
  const reel = etat.mode === "reel";
  return [barre({titre: "Réglages", retour: "#/"}), ...bandeaux(),
    h("main", {class: "contenu reglages"},
      reel ? h("h2", {class: "titre-groupe"}, "OneDrive") : null,
      reel ? h("ul", {class: "liste"},
        ligne("Compte", jetons ? (etat.compte ? `${etat.compte.nom}${etat.compte.courriel ? " (" + etat.compte.courriel + ")" : ""}` : "Connecté")
          : etat.compte ? "Reconnexion nécessaire" : "Non connecté"),
        ligne("Dossier", etat.dossier ? "OneDrive › " + etat.dossier.chemin : "Aucun"),
        ligne("Dernière synchronisation", etat.synchro ? quand(etat.synchro) : "Jamais"),
        ligne("Modifications", etat.modifie ? "Pas encore envoyées" : "Toutes envoyées")) : null,
      reel ? h("div", {class: "boutons"},
        jetons ? h("button", {class: "bouton principal", onclick: () => synchroniser(true)}, "Synchroniser maintenant")
          : h("button", {class: "bouton principal", disabled: !CONFIG.clientId, onclick: () => connecter(false)}, "Se connecter avec Microsoft"),
        jetons ? h("button", {class: "bouton", onclick: () => aller("#/dossier")}, "Changer de dossier") : null,
        jetons ? h("button", {class: "bouton danger", onclick: seDeconnecter}, "Se déconnecter") : null) : null,
      h("h2", {class: "titre-groupe"}, "Démonstration"),
      h("p", {class: "note-info"}, reel ? "Un agenda d'exemple, gardé à part sur l'iPhone : il n'est jamais envoyé dans OneDrive."
        : "Vous êtes dans la démonstration. Votre agenda OneDrive n'est pas touché."),
      h("div", {class: "boutons"}, reel
        ? h("button", {class: "bouton", onclick: () => ouvrirDemo(false)}, "Ouvrir la démonstration")
        : [h("button", {class: "bouton principal", onclick: () => changerMode("reel").then(() => aller("#/"))}, "Quitter la démonstration"),
          h("button", {class: "bouton", onclick: () => ouvrirDemo(true)}, "Recommencer la démonstration")]),
      h("h2", {class: "titre-groupe"}, "À propos"),
      h("ul", {class: "liste"},
        ligne("Version de l'app", VERSION_APP),
        ligne("Format des données", "version " + VERSION_DONNEES + " (Cahier à spirale 1.11)"),
        ligne("Adresse de retour (Entra)", adresseRetour())),
      reel && etat.cahier ? h("div", {class: "boutons"}, h("button", {class: "bouton danger", onclick: effacerCopie}, "Effacer la copie de l'iPhone")) : null)];
}
async function seDeconnecter() {
  if (etat.modifie && !(await confirmer("Des modifications ne sont pas encore envoyées dans OneDrive. Elles restent sur l'iPhone et partiront à la prochaine connexion.\n\nSe déconnecter quand même ?", "Se déconnecter", false))) return;
  jetons = null;
  await BD.ecrire("jetons", null);
  rendre();
  avis("Déconnecté");
}
async function effacerCopie() {
  const texte = etat.modifie ? "Des modifications ne sont pas encore envoyées dans OneDrive : elles seront perdues.\n\nEffacer la copie de l'agenda sur l'iPhone ?"
    : "Effacer la copie de l'agenda sur l'iPhone ? Elle sera téléchargée de nouveau depuis OneDrive.";
  if (!(await confirmer(texte, "Effacer"))) return;
  const garder = {dossier: etat.dossier, compte: etat.compte};
  Object.assign(etat, vierge(), garder);
  await sauverLocal(true);
  aller("#/");
  synchroniser();
}
async function changerMode(mode) {
  await sauverLocal(true);
  etat.mode = mode;
  await BD.ecrire("mode", mode);
  Object.assign(etat, vierge(), (await BD.lire("etat-" + mode)) || {});
  _feries.clear();
  semaines.clear();
  if (mode === "reel") planifierSynchro(200);
}
/** Les dates de la démonstration suivent le jour où elle est ouverte (décalage en semaines entières). */
function decalerDemo(c, jour) {
  const ref = dateIso(c._demo_reference);
  delete c._demo_reference;
  if (!ref) return c;
  const n = Math.round((jour - ref) / (7 * 86400000)) * 7;
  if (!n) return c;
  const decaler = (t) => (dateIso(t) ? iso(ajouterJours(dateIso(t), n)) : t);
  for (const s of c.sections) for (const p of s.pages) {
    if (p.jours) p.jours = Object.fromEntries(Object.entries(p.jours).map(([k, v]) => [decaler(k), v]));
    for (const t of p.taches || []) t.echeance = t.echeance ? decaler(t.echeance) : null;
    for (const o of p.operations || []) o.date = decaler(o.date);
  }
  return c;
}
async function ouvrirDemo(recommencer) {
  await changerMode("demo");
  if (recommencer || !etat.cahier) {
    try {
      const r = await fetch("demo.json", {cache: "no-cache"});
      etat.cahier = decalerDemo(await r.json(), aujourdhui());
      etat.modifie = false;
      await sauverLocal(true);
    } catch (e) {
      await changerMode("reel");
      avis("Démonstration indisponible hors ligne (ouvrez l'app une fois avec le réseau)");
      return;
    }
  }
  aller("#/");
}

/* ----- choix du dossier OneDrive ----- */
const nav = {pile: null, enfants: null, propose: undefined, erreur: null, jeton: 0};
function ecranDossier() {
  if (!jetons) { setTimeout(() => aller("#/"), 0); return [barre({titre: "Dossier"})]; }
  if (!nav.pile) { nav.pile = [{id: null, nom: "OneDrive"}]; nav.enfants = null; nav.propose = undefined; chargerDossier(); proposerDossier(); }
  const ici = nav.pile[nav.pile.length - 1];
  const chemin = nav.pile.slice(1).map((x) => x.nom).join("/");
  const contientCahier = (nav.enfants || []).some((e) => e.file && e.name.toLowerCase() === FICHIER);
  return [barre({titre: "Dossier de l'agenda", retour: etat.cahier ? "#/reglages" : null}), ...bandeaux(),
    h("main", {class: "contenu"},
      h("p", {class: "note-info", style: {marginTop: 0}}, "Choisissez le dossier de OneDrive qui contient cahier.json (le dossier de l'agenda choisi sur le PC)."),
      nav.propose ? h("div", {class: "trouve"}, h("b", {}, `Agenda trouvé : OneDrive › ${nav.propose.chemin}`),
        h("p", {class: "note-info"}, nav.propose.detail),
        h("button", {class: "bouton principal large", onclick: () => choisirDossier(nav.propose)}, "Utiliser ce dossier")) : null,
      h("h2", {class: "titre-groupe"}, "Parcourir OneDrive"),
      h("div", {class: "fil"}, nav.pile.map((x, i) => [i ? "›" : null, h("button", {onclick: () => { nav.pile = nav.pile.slice(0, i + 1); chargerDossier(); }}, x.nom)])),
      nav.erreur ? h("div", {class: "bandeau erreur"}, nav.erreur) : null,
      h("ul", {class: "liste"}, nav.enfants === null ? h("li", {class: "vide"}, "Chargement…")
        : nav.enfants.filter((e) => e.folder).length ? nav.enfants.filter((e) => e.folder).map((e) => h("li", {}, h("button", {class: "ligne",
          onclick: () => { nav.pile.push({id: e.id, nom: e.name}); chargerDossier(); }},
        h("span", {class: "libelle"}, h("b", {}, e.name), h("small", {}, `${e.folder.childCount} élément${e.folder.childCount > 1 ? "s" : ""}`)), h("span", {class: "chevron"}, "›"))))
          : h("li", {class: "vide"}, "Aucun sous-dossier")),
      ici.id ? h("div", {class: "boutons"}, h("button", {class: "bouton principal large", disabled: !contientCahier,
        onclick: () => choisirDossier({id: ici.id, chemin})}, contientCahier ? `Utiliser « ${ici.nom} » (cahier.json trouvé)` : "Pas de cahier.json dans ce dossier")) : null)];
}
async function chargerDossier() {
  const jeton_ = ++nav.jeton;
  nav.enfants = null;
  nav.erreur = null;
  rendre();
  const ici = nav.pile[nav.pile.length - 1];
  try {
    const d = await graphJson((ici.id ? `/me/drive/items/${ici.id}` : "/me/drive/root") + "/children?$top=200&$select=id,name,folder,file");
    if (jeton_ !== nav.jeton) return;
    nav.enfants = (d.value || []).sort((a, b) => a.name.localeCompare(b.name, "fr"));
  } catch (e) {
    if (jeton_ !== nav.jeton) return;
    nav.enfants = [];
    nav.erreur = e instanceof ErreurConnexion ? "Reconnectez-vous." : e instanceof TypeError ? "Hors ligne." : e.message;
  }
  if (route()[0] === "dossier" || !etat.dossier) rendre();
}
async function proposerDossier() {
  const chemin = String(CONFIG.dossier || "").replace(/^\/+|\/+$/g, "");
  if (!chemin) return;
  try {
    const dossier = await graphJson(`/me/drive/root:/${chemin.split("/").map(segment).join("/")}`);
    const r = await graph(`/me/drive/items/${dossier.id}:/${FICHIER}`);
    if (!r.ok) return;
    const f = await r.json();
    const qui = f.lastModifiedBy && f.lastModifiedBy.user && f.lastModifiedBy.user.displayName;
    nav.propose = {id: dossier.id, chemin, detail: `cahier.json modifié ${quand(f.lastModifiedDateTime)}${qui ? " par " + qui : ""}.`};
    if (route()[0] === "dossier" || !etat.dossier) rendre();
  } catch (e) { /* pas de dossier proposé */ }
}
async function choisirDossier(d) {
  if (etat.dossier && etat.dossier.id !== d.id && etat.modifie
    && !(await confirmer("Des modifications de l'agenda actuel ne sont pas encore envoyées : elles seront perdues.\n\nChanger de dossier ?", "Changer"))) return;
  if (!etat.dossier || etat.dossier.id !== d.id) Object.assign(etat, vierge(), {compte: etat.compte});
  etat.dossier = {id: d.id, chemin: d.chemin};
  nav.pile = null;
  await sauverLocal(true);
  aller("#/");
  synchroniser();
}

/* ===================================================================== démarrage */

async function demarrer() {
  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    // Nouvelle version publiée : le service worker l'installe, puis l'app se recharge d'elle-même (pas pendant une
    // saisie ; la copie locale est enregistrée avant). Au premier lancement, rien à recharger.
    const avaitControleur = !!navigator.serviceWorker.controller;
    let rechargement = false;
    navigator.serviceWorker.addEventListener("controllerchange", () => {
      if (!avaitControleur || rechargement) return;
      rechargement = true;
      const recharger = () => {
        if (enSaisie() || feuilleOuverte()) { setTimeout(recharger, 2000); return; }
        sauverLocal(true).then(() => location.reload());
      };
      recharger();
    });
    navigator.serviceWorker.register("sw.js").then((reg) => {
      document.addEventListener("visibilitychange", () => { if (!document.hidden) reg.update().catch(() => {}); });
    }).catch(() => {});
  }
  try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) { /* facultatif */ }
  await chargerEtat();
  const retour = new URLSearchParams(location.search).has("code") || new URLSearchParams(location.search).has("error");
  try { await retourConnexion(); } catch (e) { sync.erreur = "Connexion impossible : " + (e.message || e); }
  // À l'ouverture, la couverture de l'agenda (comme sur le PC) ; un toucher ouvre l'agenda.
  if (etat.cahier && !retour && ["", "#", "#/"].includes(location.hash)) history.replaceState(null, "", "#/couverture");
  window.addEventListener("hashchange", rendre);
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) {
      cacheDepuis = Date.now();
      sauverLocal(true);
      return;
    }
    // Retour dans l'app après une absence : la couverture, comme à l'ouverture ; un toucher reprend où l'on était.
    if (cacheDepuis && Date.now() - cacheDepuis >= ABSENCE_COUVERTURE && etat.cahier && !feuilleOuverte()
      && route()[0] !== "couverture") {
      apresCouverture = location.hash || "#/";
      history.replaceState(null, "", "#/couverture");
      rendre();
    }
    cacheDepuis = 0;
    planifierSynchro(300);
  });
  window.addEventListener("pagehide", () => sauverLocal(true));
  window.addEventListener("online", () => { sync.horsLigne = false; majStatut(); planifierSynchro(300); });
  window.addEventListener("offline", () => { sync.horsLigne = true; majStatut(); });
  document.addEventListener("focusout", () => setTimeout(apresSaisie, 80));
  setInterval(() => { if (!document.hidden) synchroniser(); }, 60000);
  rendre();
  synchroniser();
}

// Fonctions exposées pour tests.html et tests-synchro.html (OneDrive simulé)
window.AgendaTests = {interne: {etat, sync, synchroniser, modifier, adopter, envoyerPiece, definirJetons: (j) => { jetons = j; },
  annulerSynchro: () => clearTimeout(minuterieSynchro)},
  paques, feriesQuebec, montantTexte, lireMontant, notesSection, entreesSection, finNote, debutNote,
  sujetDe, finSujet, nouvellePage, modeleDe, ajouterPageNote, deplacerPage, supprimerPages, poserCellule,
  appliquerSaisieTexte, fusionner, soldes, largeursColonnes, policePage, debutSemaine, decalerDemo, canon, pageVide, iso, dateIso, OBJ};

if (document.body && document.body.dataset.app === "agenda") demarrer();
})();
