# Générateur de vidéos UGC (n8n + OpenAI Sora 2)

Workflow n8n qui produit automatiquement une vidéo publicitaire de style UGC
(TikTok / Reels / Shorts) à partir d'une simple photo de produit.

Fichier à importer : [`ugc-video-generator.json`](./ugc-video-generator.json)

## Fonctionnement

```
Formulaire ─► Normaliser ─┬─► Archiver la photo (Drive)
                          └─► Agent IA (GPT-4o + vision) ─► Créer la vidéo (Sora 2)
                                                              │
                     ┌──────────── Attendre 60 s ◄────────────┘
                     ▼                    ▲
              Vérifier le statut ─► Vidéo prête ? ── non ─► Échec / délai ? ── non ─┘
                                        │ oui                     │ oui
                                        ▼                         ▼
                           Télécharger la vidéo           Arrêter avec erreur
                                        ▼
                           Enregistrer la vidéo (Drive) ─► Envoyer par e-mail (Gmail)
```

1. **Formulaire produit** : nom, description, public cible, langue, e-mail et photo.
2. **Agent IA** : GPT-4o analyse la photo et renvoie, via le *Structured Output
   Parser*, une description fidèle du produit, une accroche, un script de 8 s,
   le prompt Sora (en anglais), une légende et des hashtags.
3. **Sora 2** (`POST /v1/videos`) génère une vidéo verticale 720×1280 de 8 s.
4. Le workflow interroge le statut toutes les 60 s (20 tentatives maximum).
5. La vidéo est téléchargée, enregistrée sur Google Drive, puis le lien, le
   script et la légende sont envoyés par e-mail.

## Installation

1. Dans n8n : **Workflows → Import from File** et choisissez `ugc-video-generator.json`.
2. Créez ou sélectionnez les identifiants :
   - **OpenAI API** sur *OpenAI Chat Model*, *Créer la vidéo*, *Vérifier le statut*
     et *Télécharger la vidéo*. Le compte doit avoir accès à l'API Sora 2.
   - **Google Drive OAuth2** sur les deux nœuds Drive (changez le dossier cible
     si besoin, par défaut la racine de « My Drive »).
   - **Gmail OAuth2** sur *Envoyer par e-mail*.
3. Activez le workflow, puis ouvrez l'URL de production du nœud *Formulaire produit*.

## Réglages utiles

| Paramètre | Où | Valeur par défaut |
|---|---|---|
| Modèle vidéo | *Créer la vidéo* → `model` | `sora-2` (`sora-2-pro` pour plus de qualité) |
| Durée | *Créer la vidéo* → `seconds` | `8` (`4`, `8` ou `12`) |
| Format | *Créer la vidéo* → `size` | `720x1280` (vertical) |
| Intervalle de vérification | *Attendre 60 s* | 60 s |
| Nombre maximal de vérifications | *Échec ou délai dépassé ?* → `$runIndex >= 20` | 20 |
| Style et consignes du script | *Agent IA* → System Message | — |

## Remarques

- La génération est facturée par seconde de vidéo par OpenAI : consultez la
  grille tarifaire avant de partager le formulaire publiquement.
- Sora refuse les personnes réelles et les célébrités : le prompt demande
  toujours une personne fictive.
- La photo n'est pas envoyée à Sora : l'agent décrit le produit en détail dans le
  prompt. Pour un rendu plus fidèle, vous pouvez transmettre la photo comme
  `input_reference` (requête multipart), à condition qu'elle ait exactement les
  dimensions de la vidéo (720×1280).
