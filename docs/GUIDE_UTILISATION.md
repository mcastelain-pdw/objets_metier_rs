# Guide d'utilisation des objets COM Sage 100 en Rust

## 🎯 Vue d'ensemble

Ce projet montre comment utiliser la bibliothèque COM **objets100c.dll** de Sage 100 depuis Rust.

## 📋 Prérequis

1. **Rust** installé
2. **objets100c.dll** disponible dans votre projet
3. **Privilèges administrateur** pour enregistrer la DLL

## 🔧 Configuration

### 1. Cargo.toml

```toml
[dependencies]
windows = { version = "0.52", features = [
    "Win32_System_Com",
    "Win32_Foundation",
    "Win32_System_Registry",
    "Win32_System_Variant"
]}
```

### 2. Enregistrement de la DLL

```powershell
# En tant qu'administrateur
regsvr32 "chemin\vers\objets100c.dll"
```

## 🎯 CLSID disponibles

Après enregistrement, les CLSID suivants sont disponibles :

| CLSID | ProgID | Description |
|-------|---------|-------------|
| `309DE0FB-9FB8-4F4E-8295-CC60C60DAA33` | `Objets100c.Cpta.Stream.1` | CBSCptaApplication100c (Comptabilité) |
| `ED0EC116-16B8-44CC-A68A-41BF6E15EB3F` | `Objets100c.Cial.Stream.1` | CBSCialApplication100C (Commercial) |

## 📝 Exemple simple d'utilisation

```rust
use windows::{
    core::*,
    Win32::System::Com::*,
};

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

fn main() -> Result<()> {
    unsafe {
        // 1. Initialiser COM
        CoInitializeEx(None, COINIT_APARTMENTTHREADED)?;
        
        // 2. Créer l'instance COM
        let clsid = parse_clsid(BSCPTA_CLSID)?;
        let instance: IUnknown = CoCreateInstance(&clsid, None, CLSCTX_INPROC_SERVER)?;
        let dispatch: IDispatch = instance.cast()?;
        
        println!("✅ Instance BSCPTAApplication100c créée !");
        
        // 3. Utiliser l'objet (voir méthodes ci-dessous)
        
        // 4. Nettoyer
        CoUninitialize();
    }
    
    Ok(())
}

unsafe fn parse_clsid(clsid_str: &str) -> Result<GUID> {
    let clsid_wide: Vec<u16> = format!("{{{}}}", clsid_str)
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();
    
    CLSIDFromString(PCWSTR(clsid_wide.as_ptr()))
}
```

## 🔧 Méthodes disponibles (IBSCPTAApplication3)

| ID | Méthode | Description |
|----|---------|-------------|
| 1 | `IsOpen` | Vérifie si une base est ouverte |
| 2 | `Open` | Ouvre une base de données |
| 3 | `Close` | Ferme la base de données |
| 4 | `Create` | Crée une nouvelle base |
| 5 | `DatabaseInfo` | Informations sur la base |
| 6 | `Synchro` | Synchronisation |
| 7 | `ReadFrom` | Lecture depuis |
| 10 | `Name` | Nom de la base |

## 🚀 Exemples d'appel de méthodes

### Méthode simple (sans paramètres)

```rust
unsafe fn call_simple_method(dispatch: &IDispatch, method_id: i32) -> Result<()> {
    let mut result = VARIANT::default();
    let mut excep_info = EXCEPINFO::default();
    let mut arg_err: u32 = 0;
    
    let dispparams = DISPPARAMS {
        rgvarg: std::ptr::null_mut(),
        rgdispidNamedArgs: std::ptr::null_mut(),
        cArgs: 0,
        cNamedArgs: 0,
    };
    
    dispatch.Invoke(
        method_id,
        &GUID::zeroed(),
        0,
        DISPATCH_METHOD | DISPATCH_PROPERTYGET,
        &dispparams,
        Some(&mut result),
        Some(&mut excep_info),
        Some(&mut arg_err),
    )?;
    
    Ok(())
}

// Utilisation
call_simple_method(&dispatch, 1)?; // IsOpen()
call_simple_method(&dispatch, 10)?; // Name
```

### Méthode Open() avec paramètres

Pour la méthode `Open()`, vous devrez généralement passer :
- Le chemin vers le fichier de base de données
- Les paramètres de connexion (utilisateur, mot de passe)
- D'autres options selon la documentation Sage

```rust
// Exemple conceptuel (à adapter selon la documentation Sage)
unsafe fn open_database(dispatch: &IDispatch, database_path: &str) -> Result<()> {
    // Créer un VARIANT avec le chemin de la base
    let path_variant = create_bstr_variant(database_path)?;
    let params = [path_variant];
    
    // Appeler Open() avec les paramètres
    call_method_with_params(dispatch, 2, &params)
}
```

## 🔍 Outils de découverte

### Découvrir les CLSID

```bash
cargo run --bin discover_clsid_fixed
```

### Test simple

```bash
cargo run --bin bscpta_simple
```

### Test avancé

```bash
cargo run --bin bscpta_advanced_v2
```

## ⚠️ Notes importantes

1. **Architecture** : Assurez-vous que votre application Rust et la DLL ont la même architecture (32bit/64bit)

2. **Privilèges** : L'enregistrement de la DLL nécessite des privilèges administrateur

3. **Documentation** : Consultez la documentation officielle Sage pour :
   - Les paramètres exacts des méthodes
   - Les valeurs de retour attendues
   - Les codes d'erreur spécifiques

4. **Gestion d'erreurs** : Utilisez toujours des blocs `try-catch` et vérifiez les codes d'erreur COM

## 🛠️ Dépannage

### Erreur "Classe non enregistrée" (0x80040154)

```bash
# Réenregistrer la DLL
regsvr32 "chemin\vers\objets100c.dll"
```

### Problèmes d'architecture

Vérifiez que votre exécutable Rust et la DLL ont la même architecture :
- Utilisez `cargo build --target i686-pc-windows-msvc` pour 32bit
- Utilisez `cargo build --target x86_64-pc-windows-msvc` pour 64bit

### Erreurs de méthodes

1. Vérifiez les ID de méthodes avec l'outil de découverte
2. Consultez la documentation Sage pour les paramètres requis
3. Testez d'abord avec des méthodes simples (IsOpen, Name)

## 📚 Ressources

- [Documentation Windows COM en Rust](https://docs.rs/windows/)
- [Documentation officielle Sage 100]
- [Exemples dans ce projet](./src/)

## ✅ Checklist de démarrage

- [ ] DLL copiée dans le projet
- [ ] DLL enregistrée avec regsvr32
- [ ] Dépendances Cargo.toml configurées
- [ ] Test de découverte CLSID réussi
- [ ] Exemple simple fonctionnel
- [ ] Documentation Sage consultée pour les méthodes spécifiques
