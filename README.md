# Gestionnaire d'Accès Google Sheets

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

Une solution de gouvernance centralisée pour gérer, sécuriser et auditer l'accès aux feuilles de calcul Google Sheets au sein de votre organisation.

## 📋 Présentation

Ce projet est une **Web App Google Apps Script** conçue pour résoudre les problèmes de conflits de versions et d'accès non autorisés sur les fichiers partagés. Il fonctionne sur un principe de **Check-in / Check-out** (comme une bibliothèque) :

1. Les utilisateurs se connectent à un portail web sécurisé.
2. Ils demandent l'accès à une feuille spécifique.
3. Le système vérifie si la feuille est libre, puis accorde les droits d'édition (ou de lecture) *uniquement* à cet utilisateur.
4. Une fois le travail terminé, l'utilisateur "rend" la feuille, révoquant ses propres droits.

## ✨ Fonctionnalités clés

### Pour les utilisateurs
* **Interface Web Moderne :** Tableau de bord intuitif, responsive, avec gestion de thèmes (Classique, Océan, Forêt).
* **Accès Immédiat :** Attribution automatique des permissions Google Drive (Viewer/Editor) sans intervention humaine.
* **Sécurité :** Inscription via compte Gmail et authentification par mot de passe haché.

### Pour les administrateurs
* **Contrôle Total :** Gestion centralisée des utilisateurs et des feuilles de calcul.
* **Journal d'Activité (Logs) :** Traçabilité complète des actions (qui a accédé à quoi et quand).
* **Sécurité Automatisée :** Un script de nettoyage nocturne (« cron job ») révoque tous les accès et déconnecte les utilisateurs chaque nuit à minuit pour éviter les oublis.
* **Zéro Infrastructure :** Aucune base de données externe requise (utilise `PropertiesService` de Google).

## 🛠️ Architecture technique

* **Backend :** Google Apps Script (Moteur V8).
* **Frontend :** HTML5, CSS3, JavaScript (servi via `HtmlService`).
* **Base de Données :** `PropertiesService` (ScriptProperties) pour stocker les utilisateurs, la configuration des feuilles et les logs au format JSON.
* **Sécurité :**
    * Mots de passe hachés (SHA-256).
    * Tokens de session uniques (UUID).
    * Utilisation de `LockService` pour gérer la concurrence (accès simultanés).

## 🚀 Installation et déploiement

### Prérequis
* Un compte Google (Gmail ou Workspace).
* Accès à Google Drive.

### Étapes d'installation

1.  **Création du Script :**
    * Créez un nouveau projet sur [script.google.com](https://script.google.com).
    * Copiez le contenu de `Code.gs` dans l'éditeur.
    * Créez les fichiers HTML (`index.html`, `documentation.html`) et copiez leurs contenus respectifs.

2.  **Configuration Initiale :**
    * Dans `Code.gs`, modifiez les constantes au début du fichier :
        ```javascript
        const EMAIL_SUPER_ADMIN = 'votre-email@gmail.com';
        const EMAIL_CONTACT_ADMIN = 'votre-email-admin@gmail.com';
        ```
    * Exécutez la fonction `installerDeclencheurNettoyage()` une fois depuis l'éditeur pour activer le nettoyage automatique de minuit.

3.  **Ajout de feuilles de calcul :**
    * Le système est vide au départ. Utilisez la fonction utilitaire `adminAjouterFeuille()` dans `Code.gs` pour ajouter vos premiers fichiers :
        ```javascript
        function initialiserDonnees() {
          adminAjouterFeuille('ID_DE_LA_GOOGLE_SHEET', 'Nom Affiché', 'Editeur');
        }
        ```

4.  **Déploiement :**
    * Cliquez sur **Déployer** > **Nouveau déploiement**.
    * Type : **Application Web**.
    * Exécuter en tant que : **Moi** (l'administrateur/propriétaire).
    * Qui a accès : **Tout le monde** (ou "Toute personne disposant d'un compte Google").

## 📖 Utilisation

1.  Partagez l'URL de l'application web à vos utilisateurs.
2.  Le premier utilisateur doit s'inscrire via le bouton "S'inscrire".
3.  L'administrateur valide manuellement l'inscription (actuellement via modification du JSON dans les propriétés, ou via une future interface admin).
4.  Une fois connecté, l'utilisateur sélectionne une feuille et clique sur "Ouvrir".

## 🔒 Sécurité et confidentialité

Ce projet respecte la confidentialité des données :
* Le code s'exécute entièrement dans l'environnement Google de l'utilisateur qui déploie.
* Aucune donnée ne transite par des serveurs tiers.
* Les mots de passe ne sont jamais stockés en clair.

## 📄 Licence

Distribué sous la licence MIT. Voir le fichier `LICENSE` pour plus d'informations.

---
**Développé par [Fabrice Faucheux]**
