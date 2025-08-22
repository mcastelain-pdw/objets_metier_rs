# Guide d'utilisation des objets COM Sage 100 en Rust

## 🎯 Vue d'ensemble

Ce projet fournit une **bibliothèque Rust** moderne et sécurisée pour utiliser les objets métier COM **objets100c.dll** de Sage 100c, avec classification intelligente des méthodes et propriétés.

## 📋 Prérequis

1. **Rust** installé (édition 2024 recommandée)
2. **Sage 100c** installé avec **objets100c.dll**
3. **Privilèges administrateur** pour enregistrer la DLL
4. **Windows** (COM natif requis)

## 🔧 Configuration rapide

### 1. Dépendances Cargo.toml

```toml
[dependencies]
objets_metier_rs = "0.1.2"
# Ou depuis le dépôt local :
# objets_metier_rs = { path = "../objets_metier_rs" }
```

### 2. Enregistrement de la DLL Sage

```powershell
# En tant qu'administrateur
regsvr32 "C:\Sage\Sage100c\objets100c.dll"
```

## 🚀 Démarrage rapide avec la nouvelle API

```rust
use objets_metier_rs::com::{ComInstance, SafeDispatch, MemberType};
use objets_metier_rs::errors::SageResult;

fn main() -> SageResult<()> {
    // 1. Créer une instance COM (gestion automatique)
    let instance = ComInstance::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;
    println!("✅ Instance BSCPTAApplication100c créée !");
    
    // 2. Découverte intelligente des membres
    let members = instance.list_members()?;
    let methods = instance.list_methods_only()?;
    let properties = instance.list_properties()?;
    
    println!("📊 Découvert : {} méthodes, {} propriétés", 
             methods.len(), properties.len());
    
    // 3. Affichage des Factory (propriétés importantes)
    for prop in properties.iter().filter(|p| p.name.starts_with("Factory")) {
        println!("🏭 {} -> {:?}", prop.name, prop.return_type);
    }
    
    // 4. Appels sécurisés
    let dispatch = SafeDispatch::new(instance.dispatch()?);
    
    match dispatch.call_method_by_name("IsOpen", &[]) {
        Ok(result) => println!("📖 Base ouverte: {}", result.type_name()),
        Err(e) => println!("❌ Erreur: {}", e),
    }
    
    Ok(())
    // 5. Nettoyage automatique (RAII)
}
```

## 🔍 Classification intelligente des membres COM

### Résultats pour BSCPTAApplication100c

La bibliothèque identifie automatiquement **47 membres** et les classifie intelligemment :

| Type | Nombre | Exemples |
|------|--------|----------|
| **🔧 Méthodes** | 7 | Open, Close, Create, DatabaseInfo, Synchro, ReadFrom |
| **📖 Propriétés** | 40 | FactoryTiers, FactoryClient, Name, IsOpen, FactoryCompteG |

### Propriétés Factory importantes

```rust
// Découverte automatique des Factory
let properties = instance.group_properties()?;

for (name, variants) in properties.iter() {
    if name.starts_with("Factory") {
        println!("🏭 {} -> Object métier Sage", name);
    }
}

// Résultats typiques :
// 🏭 FactoryTiers -> Object métier Sage (gestion des tiers)
// 🏭 FactoryClient -> Object métier Sage (gestion des clients)  
// 🏭 FactoryFournisseur -> Object métier Sage (gestion des fournisseurs)
// 🏭 FactoryCompteG -> Object métier Sage (plan comptable)
// 🏭 FactoryJournal -> Object métier Sage (journaux comptables)
// ... et 35+ autres Factory
```

### Méthodes d'action principales

```rust
// Classification automatique des méthodes
let methods = instance.list_methods_only()?;

for method in methods {
    match method.name.as_str() {
        "Open" => println!("📂 {} - Ouvrir base ({:?} params)", 
                          method.name, method.param_count),
        "Close" => println!("🔒 {} - Fermer base", method.name),
        "Create" => println!("➕ {} - Créer base", method.name),
        "IsOpen" => println!("❓ {} - État base", method.name),
        _ => println!("🔧 {} - Méthode ({:?} params)", 
                     method.name, method.param_count),
    }
}
```

## � Appels de méthodes sécurisés

### SafeDispatch - API moderne

```rust
use objets_metier_rs::com::{SafeDispatch, SafeVariant};

let dispatch = SafeDispatch::new(instance.dispatch()?);

// 1. Vérifier l'état
match dispatch.call_method_by_name("IsOpen", &[]) {
    Ok(result) => {
        println!("Base ouverte: {}", result.type_name());
        // Note: Conversion BSTR à implémenter dans v0.2.0
    },
    Err(e) => println!("Erreur: {}", e),
}

// 2. Obtenir le nom
match dispatch.call_method_by_name("Name", &[]) {
    Ok(result) => println!("Nom application: {}", result.type_name()),
    Err(e) => println!("Erreur: {}", e),
}

// 3. Appel avec paramètres (pour v0.2.0)
// let params = vec![SafeVariant::from_string("C:\\Data\\BIJOU.gcm")];
// dispatch.call_method_by_name("Open", &params)?;
```

### Gestion d'erreurs robuste

```rust
use objets_metier_rs::errors::{SageError, SageResult};

fn example_with_error_handling() -> SageResult<()> {
    let instance = ComInstance::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;
    
    match instance.supports_automation() {
        true => {
            println!("✅ Interface COM automation supportée");
            let dispatch = SafeDispatch::new(instance.dispatch()?);
            
            // Appels sécurisés avec gestion d'erreur
            match dispatch.call_method_by_name("DatabaseInfo", &[]) {
                Ok(result) => println!("Info DB: {}", result.type_name()),
                Err(SageError::MethodCall { method, id, message }) => {
                    println!("❌ Erreur méthode '{}' (ID: {}): {}", method, id, message);
                },
                Err(e) => println!("❌ Erreur générale: {}", e),
            }
        },
        false => {
            println!("❌ Interface COM automation non supportée");
            return Err(SageError::InternalError("Pas d'automation".to_string()));
        }
    }
    
    Ok(())
}
```

## 🎯 CLSID Sage 100c disponibles

| CLSID | Interface | Description | Statut |
|-------|-----------|-------------|--------|
| `309DE0FB-9FB8-4F4E-8295-CC60C60DAA33` | IBSCPTAApplication3 | Comptabilité | ✅ Testé |
| `ED0EC116-16B8-44CC-A68A-41BF6E15EB3F` | IBSCialApplication3 | Commercial | � Planifié v0.2.0 |

## �🔍 Outils de découverte et test

### Test rapide

```bash
# Test de l'API moderne
cargo run

# Résultat attendu :
# ✅ Instance BSCPTAApplication100c créée avec succès !
# 📋 Nom: IBSCPTAApplication3, Description: Interface fichier comptable
# 🔧 MÉTHODES disponibles (7 trouvées)
# 📋 PROPRIÉTÉS disponibles (40 trouvées)
```

### Tests unitaires

```bash
# Tous les tests
cargo test

# Tests spécifiques
cargo test test_clsid_parsing
cargo test test_member_classification
```

## 🚀 Roadmap et développement

### v0.1.2 (Actuelle) ✅
- ✅ Classification intelligente des membres COM
- ✅ Distinction précise méthodes vs propriétés  
- ✅ 40+ propriétés Factory découvertes
- ✅ API SafeDispatch pour appels sécurisés

### v0.1.3 (Prochaine - Priorité) 🔄
- 🔄 **Conversion VARIANT complète** (BSTR, VT_DATE, VT_CY, etc.)
- 🔄 **SafeVariant::to_string() fonctionnel** pour tous les types
- 🔄 **Paramètres de méthodes** : from_string(), from_i32(), from_bool()
- 🔄 **Arrays et collections** : Support VT_ARRAY, VT_SAFEARRAY
- 🔄 **Appels avec paramètres** : Open() fonctionnel avec chemin de base

### v0.2.0 (Future) 📋
- � Wrappers métier (Tiers, Client, Fournisseur, CompteG)
- � CRUD complet (Create, Read, Update, Delete)
- � Support Commercial (IBSCialApplication3)
- � Validation des données métier Sage

### v0.3.0 (Future) 📋
- 📋 Entités métier complètes (Écriture, Journal, Article)
- 📋 Validation des données métier
- 📋 Transactions et gestion d'erreurs avancée
- 📋 Documentation interactive

## ⚠️ Notes importantes et bonnes pratiques

### 1. Architecture et compatibilité
- **Architecture unifiée** : La bibliothèque gère automatiquement l'architecture COM
- **Gestion mémoire** : Pattern RAII automatique (pas de fuites mémoire)
- **Thread safety** : Utilisez `COINIT_APARTMENTTHREADED` (par défaut)

### 2. Sécurité et gestion d'erreurs
```rust
// ✅ Bonne pratique : gestion d'erreur structurée
use objets_metier_rs::errors::{SageError, SageResult};

fn safe_sage_operation() -> SageResult<String> {
    let instance = ComInstance::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;
    
    if !instance.supports_automation() {
        return Err(SageError::InternalError("Automation non supportée".to_string()));
    }
    
    let dispatch = SafeDispatch::new(instance.dispatch()?);
    let result = dispatch.call_method_by_name("Name", &[])?;
    
    Ok(result.type_name())
    // Nettoyage automatique
}

// ❌ À éviter : gestion manuelle unsafe
// unsafe { CoInitializeEx(...); /* code */ CoUninitialize(); }
```

### 3. Performance et optimisation
- **Réutilisation d'instance** : Créez une instance ComInstance et réutilisez-la
- **Batch operations** : Groupez les appels COM pour de meilleures performances
- **Lazy loading** : Chargez les Factory seulement quand nécessaire

### 4. Versions et migration
- **v0.1.x** : API stable pour la découverte et appels basiques
- **v0.2.x** : Wrappers métier et conversion VARIANT complète
- **Migration** : Backwards compatible, nouvelles fonctionnalités additives

## 🛠️ Dépannage courant

### ❌ Erreur "Classe non enregistrée" (0x80040154)

```powershell
# Solution : Réenregistrer la DLL avec privilèges admin
regsvr32 "C:\Sage\Sage100c\objets100c.dll"

# Vérification
cargo run  # Doit afficher "✅ Instance BSCPTAApplication100c créée avec succès !"
```

### ❌ Erreur "Interface non supportée"

```rust
// Vérifiez le support automation
let instance = ComInstance::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;
println!("Support automation: {}", instance.supports_automation());

// Si false, vérifiez la version de objets100c.dll
```

### ❌ Méthodes non trouvées

```rust
// Découvrez les méthodes disponibles
let methods = instance.list_methods_only()?;
for method in methods {
    println!("Méthode disponible: {} (ID: {})", method.name, method.id);
}

// Utilisez les noms exacts découverts
```

### ❌ Conversion VARIANT échoue (v0.1.2)

```rust
// v0.1.2 : Types supportés limités - SERA CORRIGÉ dans v0.1.3
match dispatch.call_method_by_name("IsOpen", &[]) {
    Ok(result) => {
        println!("Type reçu: {}", result.type_name()); // "BStr"
        // result.to_string() → "Conversion VARIANT non implémentée (BStr)"
        
        // ⚠️ WORKAROUND temporaire v0.1.2 : utiliser type_name()
        match result.type_name() {
            "BStr" => println!("Valeur BSTR reçue (conversion en v0.1.3)"),
            "Bool" => println!("Valeur booléenne reçue"),
            _ => println!("Type: {}", result.type_name()),
        }
    },
    Err(e) => println!("Erreur: {}", e),
}

// 🎯 v0.1.3 OBJECTIF : Conversion complète
// let name = dispatch.call_method_by_name("Name", &[])?.to_string()?;
// let is_open = dispatch.call_method_by_name("IsOpen", &[])?.to_bool()?;
// let params = vec![SafeVariant::from_string("C:\\Data\\BIJOU.gcm")];
// dispatch.call_method_by_name("Open", &params)?;
```

### 🎯 Priorité v0.1.3 : Pourquoi d'abord la conversion VARIANT ?

La conversion VARIANT complète est **cruciale** avant les wrappers métier car :

1. **Fondation nécessaire** : Tous les appels COM retournent des VARIANT
2. **Paramètres de méthodes** : Open(), Create() ont besoin de paramètres convertis  
3. **Valeurs de retour** : Name, IsOpen doivent retourner des types Rust
4. **Appels fonctionnels** : Actuellement limités aux méthodes sans paramètres
5. **Base pour v0.2.0** : Les wrappers métier dépendent des conversions

```rust
// v0.1.2 : Limité
let result = dispatch.call_method_by_name("IsOpen", &[])?;
// result.to_string() → Erreur

// v0.1.3 : Complet  
let is_open: bool = dispatch.call_method_by_name("IsOpen", &[])?.to_bool()?;
let name: String = dispatch.call_method_by_name("Name", &[])?.to_string()?;

// Avec paramètres
let db_path = SafeVariant::from_string("C:\\Data\\BIJOU.gcm");
dispatch.call_method_by_name("Open", &[db_path])?;

// Base solide pour v0.2.0 wrappers métier
let factory_tiers = dispatch.call_method_by_name("FactoryTiers", &[])?.to_dispatch()?;
let tiers = TiersWrapper::new(factory_tiers); // v0.2.0
```

## 📚 Ressources et aide

### Documentation
- **README.md** : Vue d'ensemble et exemples
- **CHANGELOG.md** : Historique des versions
- **Code source** : `/src/` avec documentation inline
- **Tests** : `/tests/` pour exemples d'usage

### Communauté et support
- **Issues GitHub** : Pour bugs et demandes de fonctionnalités
- **Discussions** : Pour questions et aide
- **Contributions** : PRs bienvenues !

### Documentation Sage officielle
- Consultez la documentation Sage 100c pour :
  - Paramètres exacts des méthodes Factory
  - Valeurs de retour des propriétés
  - Codes d'erreur spécifiques à Sage

## ✅ Checklist de démarrage rapide

- [ ] ✅ **Sage 100c installé** avec objets100c.dll
- [ ] ✅ **DLL enregistrée** : `regsvr32 "C:\Sage\Sage100c\objets100c.dll"`  
- [ ] ✅ **Dépendance ajoutée** : `objets_metier_rs = "0.1"`
- [ ] ✅ **Test basique** : `cargo run` affiche "✅ Instance créée"
- [ ] ✅ **Découverte** : 7 méthodes et 40 propriétés trouvées
- [ ] ✅ **Premier appel** : `IsOpen` et `Name` fonctionnent
- [ ] 🔄 **Prêt pour v0.2.0** : Wrappers métier et conversions VARIANT

## 🎯 Prochaines étapes

1. **Testez la découverte** : Explorez les 40+ propriétés Factory
2. **Contribuez** : Implémentez la conversion BSTR pour v0.2.0  
3. **Documentez** : Partagez vos cas d'usage spécifiques
4. **Attendez v0.2.0** : Wrappers métier complets (Tiers, Client, etc.)
