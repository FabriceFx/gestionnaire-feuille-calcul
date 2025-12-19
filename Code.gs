/**
 * ==================================================================================
 * GESTIONNAIRE D'ACCÈS GOOGLE SHEETS
 * Auteur : Fabrice Faucheux
 * Date : 19/12/2025
 * Version : 1.1 (Nettoyé et Optimisé)
 * ==================================================================================
 */

// ==================================================================================
// 1. CONFIGURATION & CONSTANTES
// ==================================================================================

/**
 * EMAIL DU SUPER ADMINISTRATEUR
 * Ce compte ne peut être ni modifié ni supprimé par d'autres administrateurs.
 * Remplacer par votre adresse email réelle.
 */
const EMAIL_SUPER_ADMIN = 'votre-email@gmail.com';

/**
 * EMAIL DE CONTACT ADMINISTRATEUR
 * Adresse où les messages des utilisateurs seront envoyés.
 */
const EMAIL_CONTACT_ADMIN = 'votre-email-admin@gmail.com';

/**
 * NOM DU SYSTÈME
 * Nom apparaissant dans les notifications par email.
 */
const NOM_SYSTEME = 'Gestionnaire d\'Accès Google Sheets';

/**
 * NOMBRE MAXIMUM D'ENTRÉES DANS LE JOURNAL
 * Limite pour économiser le stockage PropertiesService.
 */
const MAX_ENTREES_JOURNAL = 300;

const TYPES_EVENEMENT = {
  UTILISATEUR_INSCRIT: 'UTILISATEUR_INSCRIT',
  UTILISATEUR_APPROUVE: 'UTILISATEUR_APPROUVE',
  CONNEXION_REUSSIE: 'CONNEXION_REUSSIE',
  CONNEXION_ECHOUEE: 'CONNEXION_ECHOUEE',
  DECONNEXION: 'DECONNEXION',
  FEUILLE_AJOUTEE: 'FEUILLE_AJOUTEE',
  FEUILLE_SUPPRIMEE: 'FEUILLE_SUPPRIMEE',
  FEUILLE_EMPRUNTEE: 'FEUILLE_EMPRUNTEE', // Checkout
  FEUILLE_RENDUE: 'FEUILLE_RENDUE',       // Checkin
  PERMISSION_ACCORDEE: 'PERMISSION_ACCORDEE',
  PERMISSION_RETIREE: 'PERMISSION_RETIREE',
  NETTOYAGE_SYSTEME: 'NETTOYAGE_SYSTEME',
  CONTACT_ADMIN: 'CONTACT_ADMIN'
};

// ==================================================================================
// 2. POINT D'ENTRÉE WEB APP
// ==================================================================================

/**
 * Fonction appelée lors de la visite de l'URL de la Web App.
 * Sert le fichier HTML principal.
 * @return {HtmlOutput} L'interface utilisateur.
 */
const doGet = () => {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(NOM_SYSTEME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
};

/**
 * Récupère le contenu de la documentation HTML.
 * @return {string} Le contenu HTML.
 */
const recupererHtmlDocumentation = () => {
  try {
    return HtmlService.createHtmlOutputFromFile('documentation').getContent();
  } catch (erreur) {
    console.error(`Erreur chargement doc: ${erreur}`);
    throw new Error('Impossible de charger la documentation.');
  }
};

// ==================================================================================
// 3. UTILITAIRES & SÉCURITÉ
// ==================================================================================

/**
 * Normalise une adresse email (minuscules, sans espaces).
 * @param {string} email - L'email à traiter.
 * @return {string} L'email normalisé.
 */
const normaliserEmail = (email) => {
  return (email && typeof email === 'string') ? email.toLowerCase().trim() : '';
};

/**
 * Nettoie les entrées utilisateur pour éviter les injections simples.
 * @param {string} entree - Texte brut.
 * @param {number} longueurMax - Longueur maximale autorisée.
 * @return {string} Texte nettoyé.
 */
const nettoyerEntree = (entree, longueurMax = null) => {
  const caracteresDangereux = /[=|<(\];`$]/g;
  let nettoye = (!entree || typeof entree !== 'string') ? entree : entree.replace(caracteresDangereux, '*');
  
  if (longueurMax && typeof nettoye === 'string' && nettoye.length > longueurMax) {
    nettoye = nettoye.substring(0, longueurMax);
  }
  return nettoye;
};

/**
 * Wrapper pour gérer les verrous (concurrence) afin d'éviter les conflits d'écriture.
 * @param {string} nomVerrou - Identifiant du verrou.
 * @param {Function} fonctionRappel - La fonction à exécuter.
 */
const avecVerrou = (nomVerrou, fonctionRappel) => {
  const verrou = LockService.getScriptLock();
  try {
    verrou.waitLock(10000); // Attend jusqu'à 10 secondes
    return fonctionRappel();
  } catch (e) {
    console.warn(`Erreur verrou (${nomVerrou}): ${e}`);
    throw new Error('Le système est occupé, veuillez réessayer dans un instant.');
  } finally {
    verrou.releaseLock();
  }
};

/**
 * Hachage de mot de passe (SHA-256).
 * @param {string} motDePasse - Le mot de passe en clair.
 * @return {string} Le hash hexadécimal.
 */
const hacherMotDePasse = (motDePasse) => {
  const hashBrut = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, motDePasse);
  return hashBrut.map(octet => ('0' + (octet & 0xFF).toString(16)).slice(-2)).join('');
};

/**
 * Vérifie un mot de passe.
 * @param {string} motDePasse - Mot de passe fourni.
 * @param {string} hashStocke - Hash enregistré.
 * @return {boolean} Vrai si correspondance.
 */
const verifierMotDePasse = (motDePasse, hashStocke) => hacherMotDePasse(motDePasse) === hashStocke;

/**
 * Génère un jeton de session unique.
 * @return {string} UUID.
 */
const genererJetonSession = () => Utilities.getUuid();

// ==================================================================================
// 4. GESTION DES DONNÉES (PERSISTENCE)
// ==================================================================================

/**
 * Récupère tous les utilisateurs depuis PropertiesService.
 * @return {Object} Objet contenant les utilisateurs.
 */
const recupererTousUtilisateurs = () => {
  const props = PropertiesService.getScriptProperties();
  const jsonUtilisateurs = props.getProperty('UTILISATEURS');
  return jsonUtilisateurs ? JSON.parse(jsonUtilisateurs) : {};
};

/**
 * Sauvegarde les utilisateurs de manière atomique.
 * @param {Object} utilisateurs - Objet complet des utilisateurs.
 */
const sauvegarderUtilisateurs = (utilisateurs) => {
  return avecVerrou('utilisateurs', () => {
    PropertiesService.getScriptProperties().setProperty('UTILISATEURS', JSON.stringify(utilisateurs));
    return true;
  });
};

/**
 * Récupère toutes les feuilles de calcul.
 * @return {Object} Objet des feuilles.
 */
const recupererToutesFeuilles = () => {
  const props = PropertiesService.getScriptProperties();
  const jsonFeuilles = props.getProperty('FEUILLES');
  return jsonFeuilles ? JSON.parse(jsonFeuilles) : {};
};

/**
 * Sauvegarde les feuilles de calcul de manière atomique.
 * @param {Object} feuilles - Objet complet des feuilles.
 */
const sauvegarderFeuilles = (feuilles) => {
  return avecVerrou('feuilles', () => {
    PropertiesService.getScriptProperties().setProperty('FEUILLES', JSON.stringify(feuilles));
    return true;
  });
};

// ==================================================================================
// 5. AUTHENTIFICATION & SESSION
// ==================================================================================

/**
 * Connecte un utilisateur et crée une session.
 * @param {string} email 
 * @param {string} motDePasse 
 */
const connecterUtilisateur = (email, motDePasse) => {
  const emailNorm = normaliserEmail(nettoyerEntree(email, 100));
  const mdpSecurise = nettoyerEntree(motDePasse, 100);

  if (!emailNorm || !mdpSecurise) {
    return { succes: false, message: 'Email et mot de passe requis.' };
  }

  const utilisateurs = recupererTousUtilisateurs();
  const cleUtilisateur = `user_${emailNorm}`;
  const utilisateur = utilisateurs[cleUtilisateur];

  if (!utilisateur) {
    journaliserActivite(TYPES_EVENEMENT.CONNEXION_ECHOUEE, emailNorm, 'Utilisateur introuvable');
    return { succes: false, message: 'Email ou mot de passe incorrect.' };
  }

  if (utilisateur.approbationInscription !== 'approuve') {
    return { succes: false, message: 'Votre compte est en attente d\'approbation par l\'administrateur.' };
  }

  if (!verifierMotDePasse(mdpSecurise, utilisateur.hashMdp)) {
    journaliserActivite(TYPES_EVENEMENT.CONNEXION_ECHOUEE, emailNorm, 'Mot de passe invalide');
    return { succes: false, message: 'Email ou mot de passe incorrect.' };
  }

  // Création de session
  const jetonSession = genererJetonSession();
  
  return avecVerrou('sessions', () => {
    const props = PropertiesService.getScriptProperties();
    const jsonSessions = props.getProperty('SESSIONS_ACTIVES') || '{}';
    const sessions = JSON.parse(jsonSessions);

    sessions[jetonSession] = {
      idUtilisateur: utilisateur.idUtilisateur,
      typeUtilisateur: utilisateur.typeUtilisateur,
      heureConnexion: new Date().toISOString()
    };

    props.setProperty('SESSIONS_ACTIVES', JSON.stringify(sessions));

    journaliserActivite(TYPES_EVENEMENT.CONNEXION_REUSSIE, emailNorm, 'Connexion réussie', { type: utilisateur.typeUtilisateur });

    return {
      succes: true,
      message: 'Connexion réussie',
      jetonSession: jetonSession,
      emailUtilisateur: utilisateur.idUtilisateur,
      estAdmin: utilisateur.typeUtilisateur === 'Admin'
    };
  });
};

/**
 * Vérifie la validité d'un jeton de session.
 */
const verifierJetonSession = (jetonSession) => {
  if (!jetonSession) return { valide: false, message: 'Aucun jeton fourni' };
  
  const props = PropertiesService.getScriptProperties();
  const jsonSessions = props.getProperty('SESSIONS_ACTIVES') || '{}';
  const sessions = JSON.parse(jsonSessions);
  const session = sessions[jetonSession];
  
  if (!session) return { valide: false, message: 'Session expirée ou invalide' };
  
  return {
    valide: true,
    idUtilisateur: session.idUtilisateur,
    typeUtilisateur: session.typeUtilisateur,
    estAdmin: session.typeUtilisateur === 'Admin'
  };
};

/**
 * Déconnecte l'utilisateur et libère ses ressources (feuilles empruntées).
 */
const deconnecterUtilisateur = (jetonSession) => {
  if (!jetonSession) return { succes: false, message: 'Aucun jeton' };
  
  return avecVerrou('deconnexion', () => {
    const props = PropertiesService.getScriptProperties();
    const jsonSessions = props.getProperty('SESSIONS_ACTIVES') || '{}';
    const sessions = JSON.parse(jsonSessions);
    const session = sessions[jetonSession];

    if (!session) return { succes: true, message: 'Déjà déconnecté' };

    const emailUtilisateur = session.idUtilisateur;
    delete sessions[jetonSession];
    props.setProperty('SESSIONS_ACTIVES', JSON.stringify(sessions));

    // Libérer les feuilles empruntées lors de la déconnexion
    libererFeuillesUtilisateur(emailUtilisateur);

    journaliserActivite(TYPES_EVENEMENT.DECONNEXION, emailUtilisateur, 'Déconnexion utilisateur');
    return { succes: true, message: 'Déconnexion réussie' };
  });
};

/**
 * Enregistre un nouvel utilisateur.
 */
const inscrireNouvelUtilisateur = (email, motDePasse, nomComplet) => {
  const emailNorm = normaliserEmail(nettoyerEntree(email, 100));
  const nom = nettoyerEntree(nomComplet, 100);
  const mdp = nettoyerEntree(motDePasse, 100);

  if (!emailNorm || !mdp || !nom) return { succes: false, message: 'Champs manquants.' };
  if (!emailNorm.endsWith('@gmail.com')) return { succes: false, message: 'Compte Gmail requis.' };
  
  return avecVerrou('inscription', () => {
    const utilisateurs = recupererTousUtilisateurs();
    const cleUser = `user_${emailNorm}`;

    if (utilisateurs[cleUser]) return { succes: false, message: 'Compte déjà existant.' };

    utilisateurs[cleUser] = {
      idUtilisateur: emailNorm,
      nomComplet: nom,
      hashMdp: hacherMotDePasse(mdp),
      approbationInscription: 'en_attente', // Statut par défaut
      typeUtilisateur: 'Utilisateur',
      dateInscription: new Date().toISOString(),
      feuillesAssignees: []
    };

    sauvegarderUtilisateurs(utilisateurs);
    journaliserActivite(TYPES_EVENEMENT.UTILISATEUR_INSCRIT, emailNorm, `Nouvelle inscription: ${nom}`);

    // Notification Email à l'admin (optionnel)
    try {
      MailApp.sendEmail({
        to: EMAIL_CONTACT_ADMIN,
        subject: `${NOM_SYSTEME} - Nouvelle inscription`,
        body: `Nouvel utilisateur en attente : ${nom} (${emailNorm}).`
      });
    } catch(e) {
      console.warn("Impossible d'envoyer l'email de notif: " + e.message);
    }

    return { succes: true, message: 'Inscription réussie ! En attente de validation.' };
  });
};

// ==================================================================================
// 6. GESTION DES FEUILLES DE CALCUL
// ==================================================================================

/**
 * Récupère les feuilles disponibles pour l'interface utilisateur.
 * (Fonction précédemment manquante, maintenant intégrée).
 */
const recupererFeuillesDisponibles = (jetonSession) => {
  const session = verifierJetonSession(jetonSession);
  if (!session.valide) return { succes: false, message: 'Session invalide' };

  const feuilles = recupererToutesFeuilles();
  
  // Transformation de l'objet de stockage en tableau pour le frontend
  const listeFeuilles = Object.values(feuilles).map(f => ({
    id: f.idFeuille,
    nom: f.titre,
    inUse: !!f.utilisateurActuel, // Vrai si quelqu'un l'utilise
    groupe: f.groupe
  }));
  
  return { succes: true, feuilles: listeFeuilles };
};

/**
 * Demande d'accès ("Emprunt") à une feuille.
 */
const demanderAccesFeuille = (idFeuille, jetonSession) => {
  const session = verifierJetonSession(jetonSession);
  if (!session.valide) return { succes: false, message: 'Session invalide' };

  const emailUtilisateur = session.idUtilisateur;
  
  return avecVerrou('acces_feuille', () => {
    const feuilles = recupererToutesFeuilles();
    const cleFeuille = `feuille_${idFeuille}`;
    const feuille = feuilles[cleFeuille];

    if (!feuille) return { succes: false, message: 'Feuille introuvable.' };
    
    // Vérifier si déjà utilisée par quelqu'un d'autre
    if (feuille.utilisateurActuel && feuille.utilisateurActuel !== emailUtilisateur) {
      return { 
        succes: false, 
        message: `Cette feuille est actuellement utilisée par ${feuille.utilisateurActuel}.` 
      };
    }

    // Si déjà assigné à soi-même
    if (feuille.utilisateurActuel === emailUtilisateur) {
      return {
        succes: true,
        message: 'Vous avez déjà l\'accès.',
        urlFeuille: `https://docs.google.com/spreadsheets/d/${idFeuille}`,
        typeAcces: feuille.acces
      };
    }

    try {
      const ss = SpreadsheetApp.openById(idFeuille);
      
      if (feuille.acces === 'Editeur') {
        ss.addEditor(emailUtilisateur);
      } else {
        ss.addViewer(emailUtilisateur);
      }

      feuille.utilisateurActuel = emailUtilisateur;
      feuille.heureEmprunt = new Date().toISOString();
      sauvegarderFeuilles(feuilles);

      journaliserActivite(TYPES_EVENEMENT.FEUILLE_EMPRUNTEE, emailUtilisateur, `Accès accordé: ${feuille.titre}`);
      
      return {
        succes: true,
        message: 'Accès accordé ! Ouverture...',
        urlFeuille: `https://docs.google.com/spreadsheets/d/${idFeuille}`,
        typeAcces: feuille.acces
      };

    } catch (e) {
      console.error(e);
      return { succes: false, message: 'Erreur technique lors de l\'attribution des droits Google.' };
    }
  });
};

/**
 * Libère toutes les feuilles d'un utilisateur (Usage interne).
 * Retire les permissions Editor/Viewer sur les fichiers réels.
 */
const libererFeuillesUtilisateur = (emailUtilisateur) => {
  const feuilles = recupererToutesFeuilles();
  let modifie = false;

  Object.keys(feuilles).forEach(cle => {
    const feuille = feuilles[cle];
    if (feuille.utilisateurActuel === emailUtilisateur) {
      try {
        const ss = SpreadsheetApp.openById(feuille.idFeuille);
        ss.removeEditor(emailUtilisateur);
        ss.removeViewer(emailUtilisateur);
        
        feuille.utilisateurActuel = '';
        feuille.heureEmprunt = null;
        modifie = true;
      } catch (e) {
        console.warn(`Erreur retrait droits pour ${feuille.titre}: ${e}`);
        // On continue même en cas d'erreur
      }
    }
  });

  if (modifie) sauvegarderFeuilles(feuilles);
};

// ==================================================================================
// 7. FONCTIONS ADMINISTRATEUR
// ==================================================================================

/**
 * Récupère les données du tableau de bord admin.
 */
const recupererDonneesTableauBord = (jetonSession) => {
  const session = verifierJetonSession(jetonSession);
  if (!session.valide || !session.estAdmin) {
    return { succes: false, message: 'Accès administrateur requis' };
  }

  const utilisateurs = recupererTousUtilisateurs();
  const feuilles = recupererToutesFeuilles();
  
  // Formatage pour le frontend
  const listeUtilisateurs = Object.values(utilisateurs).map(u => ({
    email: u.idUtilisateur,
    statut: u.approbationInscription,
    type: u.typeUtilisateur,
    dateInscription: u.dateInscription
  }));
  
  const listeFeuilles = Object.values(feuilles).map(f => ({
    id: f.idFeuille,
    nom: f.titre,
    utilisateurActuel: f.utilisateurActuel || null,
    niveauPermission: f.acces,
    groupe: f.groupe
  }));

  return {
    succes: true,
    utilisateurs: listeUtilisateurs,
    feuilles: listeFeuilles,
    stockage: {
      // Simulation, à affiner selon les besoins
      pourcentageUtilise: 15, 
      utilisateursActuels: listeUtilisateurs.length
    }
  };
};

/**
 * Utilitaire Admin pour ajouter manuellement une feuille au système (Setup).
 * À lancer manuellement depuis l'éditeur.
 */
const adminAjouterFeuille = (idSpreadsheet, titre, niveauAcces = 'Editeur') => {
  const feuilles = recupererToutesFeuilles();
  const cle = `feuille_${idSpreadsheet}`;
  
  feuilles[cle] = {
    idFeuille: idSpreadsheet,
    titre: titre,
    acces: niveauAcces,
    groupe: 1,
    utilisateurActuel: '',
    heureEmprunt: null
  };
  
  sauvegarderFeuilles(feuilles);
  console.log(`Feuille "${titre}" ajoutée avec succès au système.`);
};

/**
 * Fonction exemple pour initialiser la première feuille.
 */
const initialiserDonnees = () => {
  // Remplacer par un vrai ID de Google Sheet
  adminAjouterFeuille('VOTRE_ID_SPREADSHEET_ICI', 'Fichier Test', 'Editeur');
};

// ==================================================================================
// 8. JOURNALISATION
// ==================================================================================

/**
 * Ajoute une entrée au journal d'activité.
 */
const journaliserActivite = (typeEvenement, emailUtilisateur, details, donneesSup = null) => {
  avecVerrou('journal', () => {
    const props = PropertiesService.getScriptProperties();
    const jsonJournal = props.getProperty('JOURNAL_ACTIVITE') || '[]';
    let journal = JSON.parse(jsonJournal);

    journal.unshift({
      horodatage: new Date().toISOString(),
      typeEvenement,
      emailUtilisateur: emailUtilisateur || 'Système',
      details,
      donneesSup
    });

    if (journal.length > MAX_ENTREES_JOURNAL) {
      journal = journal.slice(0, MAX_ENTREES_JOURNAL);
    }

    props.setProperty('JOURNAL_ACTIVITE', JSON.stringify(journal));
  });
};

// ==================================================================================
// 9. TÂCHES AUTOMATISÉES (TRIGGERS)
// ==================================================================================

/**
 * Installe programmatiquement le déclencheur de nettoyage nocturne.
 * À exécuter UNE SEULE FOIS manuellement depuis l'éditeur.
 */
const installerDeclencheurNettoyage = () => {
  try {
    const nomFonction = 'nettoyageMinuit';
    
    // Vérification des doublons
    const declencheursExistants = ScriptApp.getProjectTriggers();
    const existeDeja = declencheursExistants.some(trigger => trigger.getHandlerFunction() === nomFonction);

    if (existeDeja) {
      console.warn(`⚠️ Le déclencheur pour "${nomFonction}" existe déjà. Installation annulée.`);
      return;
    }

    // Création du déclencheur (Quotidien, minuit)
    ScriptApp.newTrigger(nomFonction)
      .timeBased()
      .atHour(0)
      .everyDays(1)
      .inTimezone(Session.getScriptTimeZone())
      .create();
      
    console.log(`✅ Succès : Le nettoyage automatique est programmé chaque jour entre 00h00 et 01h00.`);
  } catch (erreur) {
    console.error(`❌ Erreur lors de l'installation du déclencheur : ${erreur.toString()}`);
  }
};

/**
 * Fonction exécutée automatiquement par le déclencheur.
 * Révoque tous les accès et vide les sessions chaque nuit.
 */
const nettoyageMinuit = () => {
  console.log('🌙 Début du nettoyage de minuit...');
  
  // Utilisation d'un verrou global pour cette opération critique
  avecVerrou('nettoyage_nocturne', () => {
    const props = PropertiesService.getScriptProperties();
    
    // 1. Réinitialisation des sessions
    props.setProperty('SESSIONS_ACTIVES', '{}');

    // 2. Récupération et nettoyage des feuilles
    const feuilles = recupererToutesFeuilles();
    let modificationsEffectuees = false;

    Object.keys(feuilles).forEach(cle => {
      const feuille = feuilles[cle];
      
      // Si la feuille est marquée comme empruntée
      if (feuille.utilisateurActuel) {
        try {
          const ss = SpreadsheetApp.openById(feuille.idFeuille);
          
          // Suppression des droits
          ss.removeEditor(feuille.utilisateurActuel);
          ss.removeViewer(feuille.utilisateurActuel);
          
          console.log(`♻️ Accès retiré pour ${feuille.utilisateurActuel} sur "${feuille.titre}"`);
          
          // Mise à jour des métadonnées
          feuille.utilisateurActuel = '';
          feuille.heureEmprunt = null;
          modificationsEffectuees = true;
          
        } catch (erreur) {
          console.error(`⚠️ Erreur nettoyage feuille "${feuille.titre}" (ID: ${feuille.idFeuille}): ${erreur.message}`);
        }
      }
    });

    // 3. Sauvegarde si nécessaire
    if (modificationsEffectuees) {
      sauvegarderFeuilles(feuilles);
    }
    
    // 4. Trace dans le journal
    journaliserActivite(
      TYPES_EVENEMENT.NETTOYAGE_SYSTEME, 
      'Système', 
      'Réinitialisation nocturne effectuée avec succès'
    );
  });
};
