# Changelog

Toutes les modifications notables de ce projet sont documentées dans ce fichier.

Le format est basé sur [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
et ce projet adhère à [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2025-08-22

### ✨ Ajouté
- **Architecture de base** : Structure modulaire complète du projet
- **Gestion d'erreurs robuste** : `SageError` enum avec tous les types d'erreurs Sage
- **Wrapper COM sûr** : `ComInstance` avec gestion automatique du cycle de vie
- **Appels de méthodes sécurisés** : `SafeDispatch` avec gestion d'erreurs intégrée
- **Gestion des types** : `SafeVariant` pour wrapper les VARIANT COM
- **Utilitaires de chaînes** : `SafeString` pour gérer les BSTR COM
- **Tests unitaires** : Couverture complète des modules principaux
- **Documentation** : README complet avec roadmap et exemples

### 🔧 Fonctionnalités
- ✅ **Connexion COM automatique** : Plus besoin de gérer manuellement l'initialisation
- ✅ **Détection des méthodes** : Découverte automatique des méthodes COM disponibles
- ✅ **Appels sûrs** : Wrapper type-safe pour tous les appels COM
- ✅ **Gestion mémoire** : Pattern RAII pour la libération automatique des ressources
- ✅ **Support Sage 100c** : Compatible avec BSCPTAApplication100c (comptabilité)

### 🛠️ Technique
- **Dépendances** : windows 0.52, thiserror 1.0, serde 1.0, chrono 0.4
- **Architecture** : Modules séparés pour COM, erreurs, et entités métier
- **Qualité** : Aucun avertissement de compilation, tous les tests passent
- **Performance** : Compilation optimisée en mode release

### 📋 Exemple d'utilisation
```rust
use objets_metier_rs::{ComInstance, SafeDispatch};

// Connexion automatique à Sage 100c
let instance = ComInstance::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;

// Appels de méthodes sécurisés
let dispatch = SafeDispatch::new(instance.dispatch()?);
let result = dispatch.call_method(1, "IsOpen")?;

println!("Base ouverte: {}", result.to_string()?);
```

### 🎯 Prochaines étapes (v0.2.0)
- Implémentation complète de la conversion VARIANT
- Module Comptabilité avec Tiers, Plan Comptable, Écriture, Journal
- Méthodes métier pour Open() avec paramètres
- Support complet des opérations CRUD
- Validation des données
- Tests d'intégration

---

### Légende
- ✨ Ajouté - pour les nouvelles fonctionnalités
- 🔧 Modifié - pour les changements dans les fonctionnalités existantes
- ❌ Supprimé - pour les fonctionnalités supprimées
- 🐛 Corrigé - pour les corrections de bugs
- 🔒 Sécurité - pour les corrections de vulnérabilités
