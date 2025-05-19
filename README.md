# Word SharePoint Automation

Outil d'automatisation pour la génération de documents Word à partir d'une source Excel et leur stockage sur SharePoint/OneDrive.

## Installation

### Prérequis

- Git
- Docker et Docker Compose
- Rclone

### Étapes d'installation

1. Cloner le dépôt:
   ```bash
   git clone https://github.com/votre-repo/word-sharepoint-automation.git
   cd word-sharepoint-automation
   ```

2. Lancer l'environnement Docker:
   ```bash
   sudo docker-compose up -d
   ```

## Configuration de Rclone

Rclone est utilisé pour interagir avec OneDrive et SharePoint. Voici comment le configurer:

### Installation de Rclone

```bash
curl https://rclone.org/install.sh | sudo bash
```

### Configuration pour SharePoint

1. Lancez la configuration de rclone:
   ```bash
   rclone config
   ```

2. Suivez ces étapes:
   - Appuyez sur `n` pour un nouveau remote
   - Nom: `sharepoint`
   - Sélectionnez le type `36` (Microsoft OneDrive)
   - Acceptez les paramètres par défaut jusqu'à l'authentification
   - Pour l'authentification, utilisez le navigateur web si possible, sinon utilisez l'autre méthode proposée
   - Sélectionnez `4` (SharePoint - recherchez votre site)
   - Suivez les instructions à l'écran pour terminer l'authentification

### Configuration pour OneDrive personnel

1. Lancez à nouveau la configuration:
   ```bash
   rclone config
   ```

2. Suivez les mêmes étapes mais:
   - Utilisez un nom différent, par exemple `onedrive`
   - Sélectionnez `1` (OneDrive personnel) au lieu de `4`

### Copier la configuration Rclone dans le conteneur Docker

Si vous exécutez l'application dans Docker, copiez votre configuration:

```bash
docker cp ~/.config/rclone/rclone.conf <nom_du_conteneur>:/root/.config/rclone/rclone.conf
```

## Utilisation

L'application:
1. Télécharge régulièrement un fichier Excel depuis OneDrive
2. Télécharge un modèle Word depuis SharePoint
3. Génère un document Word pour chaque ligne de l'Excel
4. Téléverse les documents générés vers SharePoint

Pour lancer l'application:
```bash
python main.py
```

Vous pouvez configurer l'intervalle d'exécution avec la variable d'environnement `INTERVAL_MINUTES`.

## Structure des fichiers

- `main.py` - Point d'entrée de l'application
- `sharepoint_utils.py` - Fonctions utilitaires pour interagir avec SharePoint/OneDrive
- `docker-compose.yml` - Configuration Docker

## Dépannage

- Si vous rencontrez des erreurs d'authentification, vérifiez que votre configuration rclone est valide
- Pour les problèmes de chemin, assurez-vous que les chemins distants sont correctement spécifiés
- Consultez les logs pour plus d'informations sur les erreurs
