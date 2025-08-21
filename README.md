# Objets Métier Sage 100c - Wrapper Rust

[![Rust](https://img.shields.io/badge/rust-1.70+-orange.svg)](https://www.rust-lang.org)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)

## 📋 Description

Ce projet propose une bibliothèque wrapper en **Rust** pour l'API COM **Objets Métier Sage 100c**. Il permet d'interagir avec les bases de données et fonctionnalités de Sage 100c depuis des applications Rust modernes.

Sage 100c fournit uniquement des exemples pour C# et Visual Basic dans sa documentation PDF. Ce projet comble cette lacune en offrant une interface Rust moderne et sûre pour les développeurs souhaitant intégrer Sage 100c dans leurs applications.

## 🎯 Objectifs du projet

### Objectifs principaux
- ✅ **Wrapper Rust sûr** pour la DLL COM `objets100c.dll`
- ✅ **Interface ergonomique** cachant la complexité COM
- ✅ **Documentation complète** avec exemples pratiques
- ✅ **Gestion d'erreurs robuste** avec types Rust idiomatiques
- ✅ **Support des opérations CRUD** sur les données Sage
- ✅ **Abstraction des modules** Comptabilité, Commercial, Paie, etc.

### Fonctionnalités cibles
- 🔗 **Connexion aux bases Sage** (ouverture/fermeture)
- 📊 **Lecture des données** (comptes, écritures, articles, clients...)
- ✏️ **Écriture des données** (création/modification d'écritures)
- 📈 **Opérations de reporting** et d'export
- 🔍 **Requêtes avancées** avec filtres
- 🔄 **Synchronisation** et opérations batch

## 🏗️ Architecture

```
objets_metier_rs/
├── src/
│   ├── lib.rs              # API publique du wrapper
│   ├── main.rs             # Exemples et tests
│   ├── com/                # Couche COM bas niveau
│   │   ├── mod.rs
│   │   ├── instance.rs     # Gestion des instances COM
│   │   └── dispatch.rs     # Appels de méthodes COM
│   ├── modules/            # Modules métier Sage
│   │   ├── mod.rs
│   │   ├── comptabilite.rs # Module Comptabilité
│   │   ├── commercial.rs   # Module Commercial
│   │   └── paie.rs        # Module Paie
│   ├── entities/           # Entités métier
│   │   ├── mod.rs
│   │   ├── compte.rs       # Comptes comptables
│   │   ├── ecriture.rs     # Écritures comptables
│   │   ├── client.rs       # Clients
│   │   └── article.rs      # Articles
│   └── errors/             # Gestion d'erreurs
│       ├── mod.rs
│       └── sage_error.rs
├── examples/               # Exemples d'utilisation
├── docs/                   # Documentation supplémentaire
├── tests/                  # Tests d'intégration
└── README.md
```

## 🚀 Démarrage rapide

### Prérequis

1. **Rust 1.70+** installé
2. **Sage 100c** installé avec `objets100c.dll`
3. **Privilèges administrateur** pour l'enregistrement DLL
4. **Windows** (COM nécessaire)

### Installation

```bash
# Cloner le projet
git clone https://github.com/votre-username/objets_metier_rs.git
cd objets_metier_rs

# Enregistrer la DLL Sage (en tant qu'administrateur)
regsvr32 "C:\Sage\Sage100c\objets100c.dll"

# Compiler et tester
cargo build
cargo run
```

### Exemple d'utilisation

```rust
use objets_metier_rs::{SageConnection, SageError};

fn main() -> Result<(), SageError> {
    // Connexion à la base Sage
    let mut sage = SageConnection::new()?;
    sage.open_database("C:\\Sage\\Data\\BIJOU.gcm", "ADMIN", "")?;
    
    // Lecture des comptes
    let comptes = sage.comptabilite().list_comptes()?;
    println!("Nombre de comptes: {}", comptes.len());
    
    // Création d'une écriture
    let ecriture = sage.comptabilite()
        .create_ecriture()?
        .journal("VT")
        .date("01/01/2024")
        .piece("FACT001")
        .compte_debit("411000", 1000.0)
        .compte_credit("701000", 1000.0)
        .save()?;
    
    println!("Écriture créée: {}", ecriture.numero());
    
    sage.close()?;
    Ok(())
}
```

## 📚 Documentation

### Structure de la documentation

- 📖 **[Guide d'utilisation](GUIDE_UTILISATION.md)** - Configuration et premiers pas
- 🔧 **[Référence API](docs/api/)** - Documentation complète des méthodes
- 💡 **[Exemples](examples/)** - Cas d'usage pratiques
- ❓ **[FAQ](docs/FAQ.md)** - Questions fréquentes
- 🔍 **[Troubleshooting](docs/troubleshooting.md)** - Résolution de problèmes

### Modules supportés

| Module | Status | Description |
|--------|--------|-------------|
| 💼 **Comptabilité** | ✅ En cours | Comptes Tiers, Plan Comptable, écritures, journaux |
| 🛒 **Commercial** | 📋 Planifié | Clients, Fournisseurs, articles, commandes |
| 💰 **Paie** | 📋 Planifié | Employés, bulletins de paie |
| 📊 **Immobilisations** | 📋 Planifié | Biens, amortissements |
| 🏦 **Trésorerie** | 📋 Planifié | Banques, échéances |

## 🛠️ Développement

### Contribuer

1. **Fork** le projet
2. Créer une **branche feature** (`git checkout -b feature/nouvelle-fonctionnalite`)
3. **Commiter** les changements (`git commit -m 'Ajout nouvelle fonctionnalité'`)
4. **Push** vers la branche (`git push origin feature/nouvelle-fonctionnalite`)
5. Ouvrir une **Pull Request**

### Tests

```bash
# Tests unitaires
cargo test

# Tests d'intégration (nécessite Sage 100c)
cargo test --test integration

# Tests avec une base de données test
SAGE_DB_PATH="C:\\Sage\\Data\\TEST.gcm" cargo test
```

### Standards de code

- **Format** : `cargo fmt`
- **Linting** : `cargo clippy`
- **Documentation** : Toutes les APIs publiques documentées
- **Tests** : Couverture > 80%

## 📦 Dépendances

### Principales

- `windows = "0.52"` - Bindings Windows COM
- `serde = "1.0"` - Sérialisation des entités
- `chrono = "0.4"` - Gestion des dates
- `thiserror = "1.0"` - Gestion d'erreurs

### Développement

- `tokio-test` - Tests asynchrones
- `mockall` - Mocking pour les tests
- `criterion` - Benchmarks

## 📋 Roadmap

### Version 0.1.0 - Fondations ✅ **TERMINÉE**
- [x] Configuration projet Rust
- [x] Connexion COM basique
- [x] Découverte CLSID et méthodes
- [x] Wrapper sûr pour les appels COM
- [x] Gestion d'erreurs Rust

### Version 0.2.0 - Module Comptabilité
- [ ] Entités Compte, Écriture, Journal
- [ ] CRUD opérations comptables
- [ ] Validation des données
- [ ] Tests d'intégration

### Version 0.3.0 - Module Commercial
- [ ] Entités Client, Article, Commande
- [ ] Gestion des stocks
- [ ] Calculs de prix et remises

### Version 1.0.0 - Production Ready
- [ ] Documentation complète
- [ ] Performances optimisées
- [ ] Support multi-threading
- [ ] Package crates.io

## ⚠️ Limitations connues

- **Windows uniquement** - Dépendance COM native
- **Architecture** - La DLL et l'executable doivent avoir la même architecture (32/64-bit)
- **Licences Sage** - Respect des termes de licence Sage 100c
- **Version Sage** - Testé sur Sage 100c v10.05

## 🤝 Support

### Canaux de support

- 🐛 **Issues GitHub** - Bugs et demandes de fonctionnalités
- 💬 **Discussions** - Questions et aide communautaire
- 📧 **Email** - Contact direct pour les entreprises

### Ressources utiles

- [Documentation Sage 100c](https://sage.fr/documentation)
- [Guide COM en Rust](https://docs.rs/windows/)
- [Exemples C# Sage](./docs/exemples-csharp/)

## 📄 Licence

Ce projet est sous licence **MIT**. Voir le fichier [LICENSE](LICENSE) pour plus de détails.

## 👥 Contributeurs

- **[Votre nom]** - *Créateur et mainteneur principal* - [@votre-github](https://github.com/votre-github)

## 🙏 Remerciements

- **Sage** pour la documentation PDF des Objets Métier
- **Microsoft** pour les bindings Rust Windows
- **Communauté Rust** pour les outils et bibliothèques

---

<div align="center">

**[🏠 Accueil](#objets-métier-sage-100c---wrapper-rust)** • 
**[📖 Documentation](GUIDE_UTILISATION.md)** • 
**[💡 Exemples](examples/)** • 
**[🐛 Issues](https://github.com/votre-username/objets_metier_rs/issues)**

</div>
