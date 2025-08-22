# Guide d'Utilisation - Objets Métier Sage 100c v0.1.3

## 🚀 Démarrage Rapide

### Prérequis
- **Rust 1.70+** 
- **Sage 100c** installé avec `objets100c.dll` enregistrée
- **Windows** (dépendance COM native)
- **Droits administrateur** pour la première utilisation

### Installation

```bash
git clone https://github.com/votre-username/objets_metier_rs.git
cd objets_metier_rs
cargo build --release
```

### Premier Test

```bash
# Test de connexion Sage (nécessite base existante)
cargo run --example sage_connection_demo
```

## 📚 API v0.1.3

### CptaApplication - Wrapper Comptabilité

```rust
use objets_metier_rs::wrappers::CptaApplication;
use objets_metier_rs::SageResult;

// CLSID BSCPTAApplication100c
const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

let app = CptaApplication::new(BSCPTA_CLSID)?;
```

#### Méthodes principales

| Méthode | Équivalent C# | Description |
|---------|---------------|-------------|
| `new(clsid)` | `new BSCPTAApplication100c()` | Créer instance |
| `name()?` | `_mCpta.Name` | Chemin base de données |
| `set_name(path)?` | `_mCpta.Name = path` | Définir chemin base |
| `is_open()?` | `_mCpta.IsOpen` | État ouverture base |
| `open()?` | `_mCpta.Open()` | Ouvrir base |
| `close()?` | `_mCpta.Close()` | Fermer base |
| `loggable()?` | `_mCpta.Loggable` | Objet authentification |

### CptaLoggable - Wrapper Authentification

```rust
let loggable = app.loggable()?;
```

#### Méthodes authentification

| Méthode | Équivalent C# | Description |
|---------|---------------|-------------|
| `user_name()?` | `_mCpta.Loggable.UserName` | Nom utilisateur |
| `set_user_name(name)?` | `_mCpta.Loggable.UserName = name` | Définir utilisateur |
| `set_user_pwd(pwd)?` | `_mCpta.Loggable.UserPwd = pwd` | Définir mot de passe |
| `is_logged()?` | `_mCpta.Loggable.IsLogged` | État connexion |
| `is_administrator()?` | `_mCpta.Loggable.IsAdministrator` | Droits admin |

## 🎯 Exemples d'Usage

### Connexion Simple

```rust
use objets_metier_rs::wrappers::CptaApplication;

fn connexion_sage() -> SageResult<()> {
    let app = CptaApplication::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;
    
    // Configuration base
    app.set_name(r"C:\Sage\Data\DEMO.MAE")?;
    
    // Authentification
    app.loggable()?.set_user_name("<Administrateur>")?;
    app.loggable()?.set_user_pwd("")?;
    
    // Connexion
    app.open()?;
    println!("✅ Connexion réussie !");
    
    // Vérifications
    if app.is_open()? {
        println!("📋 Base: {}", app.name()?);
        println!("👤 Utilisateur: {}", app.loggable()?.user_name()?);
        println!("🔐 Connecté: {}", app.loggable()?.is_logged()?);
    }
    
    app.close()?;
    Ok(())
}
```

### Gestion d'Erreurs Avancée

```rust
match app.open() {
    Ok(()) => {
        println!("🎉 Connexion réussie");
        // Traitement...
        app.close()?;
    }
    Err(SageError::ComError { hresult, message }) => {
        eprintln!("❌ Erreur COM: HRESULT={:08X}", hresult);
        eprintln!("   Message: {}", message);
        
        // Diagnostic selon HRESULT
        match hresult {
            0x800700020 => eprintln!("💡 Cause probable: Fichier base introuvable"),
            0x80004005 => eprintln!("💡 Cause probable: Credentials incorrects"),
            _ => eprintln!("💡 Consultez la documentation Sage"),
        }
    }
    Err(e) => eprintln!("❌ Autre erreur: {}", e),
}
```

### Cycle de Vie Complet

```rust
fn cycle_complet_sage() -> SageResult<()> {
    let app = CptaApplication::new("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33")?;
    
    // 1. Configuration
    app.set_name(r"D:\TMP\BIJOU.MAE")?;
    println!("📁 Base configurée: {}", app.name()?);
    
    // 2. Authentification
    let loggable = app.loggable()?;
    loggable.set_user_name("<Administrateur>")?;
    loggable.set_user_pwd("")?;
    println!("👤 Utilisateur: {}", loggable.user_name()?);
    
    // 3. Connexion
    app.open()?;
    println!("🔓 Base ouverte");
    
    // 4. Vérifications post-connexion
    assert!(app.is_open()?);
    assert!(loggable.is_logged()?);
    
    // 5. Informations système
    println!("📊 Statut:");
    println!("   - Base: {}", app.name()?);
    println!("   - Utilisateur: {}", loggable.user_name()?);
    println!("   - Admin: {}", loggable.is_administrator()?);
    
    // 6. Fermeture propre
    app.close()?;
    println!("🔒 Base fermée");
    
    Ok(())
}
```

## 🔧 Configuration et Paramètres

### Chemins de Base Courants

```rust
// Base démo Sage
app.set_name(r"C:\Sage\Sage100c\Base\BIJOU.MAE")?;

// Base production
app.set_name(r"D:\Sage\Prod\SOCIÉTÉ.MAE")?;

// Base réseau
app.set_name(r"\\serveur\Sage\Data\COMPTA.MAE")?;
```

### Utilisateurs Courants

```rust
// Administrateur Sage (défaut)
loggable.set_user_name("<Administrateur>")?;
loggable.set_user_pwd("")?;

// Utilisateur nommé
loggable.set_user_name("DUPONT")?;
loggable.set_user_pwd("motdepasse")?;

// Authentification Windows
loggable.set_user_name("")?;  // Vide pour Windows Auth
```

## 🐛 Dépannage

### Erreurs Courantes

#### `Le chemin d'accès spécifié est introuvable`
```rust
// Vérifiez le chemin
let path = r"D:\TMP\BIJOU.MAE";
if !std::path::Path::new(path).exists() {
    eprintln!("❌ Fichier inexistant: {}", path);
}
```

#### `Accès refusé` ou credentials
```rust
// Testez différents utilisateurs
let users = ["<Administrateur>", "SUPERVISEUR", ""];
for user in users {
    loggable.set_user_name(user)?;
    match app.open() {
        Ok(()) => { 
            println!("✅ Connexion réussie avec: {}", user);
            break;
        }
        Err(_) => continue,
    }
}
```

#### `Classe non enregistrée`
```powershell
# En tant qu'administrateur
regsvr32 "C:\Sage\Sage100c\objets100c.dll"
```

### Debug et Logging

```rust
// Activation du debug détaillé
use objets_metier_rs::com::SafeDispatch;

let dispatch = SafeDispatch::new(app.dispatch()?);

// Test méthodes individuelles
match dispatch.call_method_by_name("IsOpen", &[]) {
    Ok(result) => println!("IsOpen: {:?}", result),
    Err(e) => println!("Erreur IsOpen: {}", e),
}
```

## 🚀 Prochaines Étapes

### v0.2.0 - Module Commercial
```rust
// Prévu pour v0.2.0
use objets_metier_rs::wrappers::CialApplication;

let cial = CialApplication::new(CIAL_CLSID)?;
let clients = cial.factory_clients()?;
// ...
```

### Contribution

Voir [CONTRIBUTING.md](CONTRIBUTING.md) pour contribuer au projet.

### Support

- 🐛 [Issues GitHub](https://github.com/votre-username/objets_metier_rs/issues)
- 📧 Email: contact@votre-domaine.com
- 💬 Discussions: [GitHub Discussions](https://github.com/votre-username/objets_metier_rs/discussions)
