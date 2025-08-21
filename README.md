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
use objets_metier_rs::com::ComInstance;
use objets_metier_rs::errors::SageResult;

fn main() -> SageResult<()> {
    // Créer une instance COM de BSCPTAApplication100c
    let instance = ComInstance::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;
    println!("✅ Instance COM créée avec succès !");
    
    // Vérifier le support de l'automation
    if instance.supports_automation() {
        // Obtenir les informations de type
        let type_info = instance.get_type_info()?;
        println!("📋 {}", type_info);
        
        // Découvrir les méthodes disponibles
        let methods = instance.list_methods()?;
        println!("🔧 {} méthodes trouvées", methods.len());
        
        // Découvrir et séparer méthodes/propriétés (v0.1.0+)
        let members = instance.list_members()?;
        let methods_only = instance.list_methods_only()?;
        let properties = instance.group_properties()?;
        
        println!("📊 {} membres total", members.len());
        println!("🔧 {} méthodes pures", methods_only.len()); 
        println!("📋 {} groupes de propriétés", properties.len());
        
        // Appels de méthodes sécurisés
        use objets_metier_rs::com::SafeDispatch;
        let dispatch = SafeDispatch::new(instance.dispatch()?);
        
        match dispatch.call_method_by_name("IsOpen", &[]) {
            Ok(result) => println!("IsOpen: {}", result.type_name()),
            Err(e) => println!("Erreur: {}", e),
        }
    }
    
    Ok(())
    // Instance libérée automatiquement (RAII)
}
    
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

## � Découverte des interfaces COM

### Inspection intelligente des membres

La bibliothèque offre une classification intelligente des membres COM basée sur les conventions Sage 100c :

```rust
use objets_metier_rs::com::{ComInstance, MemberType};

let instance = ComInstance::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;

// Découverte avec classification intelligente
let members = instance.list_members()?;
for member in members {
    match member.member_type {
        MemberType::Method => println!("🔧 Méthode: {} ({:?} params)", 
                                      member.name, member.param_count),
        MemberType::PropertyGet => println!("📖 Propriété: {} -> {:?}", 
                                           member.name, member.return_type),
        // ...
    }
}

// Résultats typiques pour BSCPTAApplication100c:
// 🔧 7 méthodes (Open, Close, Create, DatabaseInfo, etc.)
// 📖 40 propriétés (FactoryTiers, FactoryClient, Name, IsOpen, etc.)
```

### Classification automatique

L'algorithme de classification reconnaît :

- **Factory*** → Propriétés retournant des objets métier
- **Is***, **Name**, **Version** → Propriétés d'état/information  
- **Open**, **Close**, **Create** → Méthodes d'action
- **DatabaseInfo**, **Synchro** → Méthodes de traitement

### Filtrage par type

```rust
// Filtrage avancé
let methods_only = instance.list_methods_only()?;     // 7 méthodes
let properties = instance.list_properties()?;         // 40 propriétés  
let grouped_props = instance.group_properties()?;     // Propriétés groupées

println!("Trouvé {} méthodes et {} propriétés", 
         methods_only.len(), properties.len());

// Exemples de propriétés Factory découvertes:
// - FactoryTiers -> Object (gestion des tiers)
// - FactoryClient -> Object (gestion des clients)  
// - FactoryFournisseur -> Object (gestion des fournisseurs)
// - FactoryCompteG -> Object (gestion du plan comptable)
```

### Informations des membres

Chaque membre découvert fournit :

- **ID** : Identifiant unique COM (DISPID)
- **Nom** : Nom de la méthode/propriété
- **Type** : Method, PropertyGet, PropertyPut, PropertyPutRef
- **Paramètres** : Nombre de paramètres estimé selon le type
- **Type de retour** : Type de la valeur retournée (Object, String, Boolean, etc.)

### Appels sécurisés

```rust
use objets_metier_rs::com::SafeDispatch;

let dispatch = SafeDispatch::new(instance.dispatch()?);

// Appel par nom avec gestion d'erreur
match dispatch.call_method_by_name("IsOpen", &[]) {
    Ok(result) => println!("Base ouverte: {}", result.type_name()),
    Err(e) => println!("Erreur: {}", e),
}
```

## �🛠️ Développement

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
- [ ] Entités Tiers, Plan Comptable, Écriture, Journal
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
- **Architecture** - La DLL et l'executable doivent avoir la même architecture (32-bit)
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
