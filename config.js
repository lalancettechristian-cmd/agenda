/* Réglages de l'app iPhone (voir GUIDE-IPHONE.md).
   clientId : « ID d'application (client) » de l'inscription créée dans Microsoft Entra. Tant qu'il est vide, l'app
   fonctionne seulement en mode démonstration. */
window.AGENDA_CONFIG = {
  clientId: "",
  // OneDrive personnel : « consumers ». Mettre « common » pour accepter aussi un compte professionnel ou scolaire.
  autorite: "consumers",
  // Dossier proposé dans OneDrive (celui qui contient cahier.json) ; on peut en choisir un autre dans l'app.
  dossier: "Agenda",
};
