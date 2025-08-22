use objets_metier_rs::wrappers::CptaApplication;
use objets_metier_rs::errors::SageResult;

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

fn main() -> SageResult<()> {
    println!("🚀 Sage 100c - Connexion automatique style C# v0.1.3");
    println!("═══════════════════════════════════════════════════════");
    
    // Reproduction exacte du code C# fourni par Sage
    demo_csharp_connection_style()?;
    
    Ok(())
}

/// REPRODUCTION EXACTE DU CODE C# DE SAGE
/// 
/// Code C# original :
/// ```csharp
/// private BSCPTAApplication100c _mCpta;
/// _mCpta.Name = "D:\\TMP\\BIJOU.MAE";
/// _mCpta.Loggable.UserName = "<Administrateur>";
/// _mCpta.Loggable.UserPwd = "";
/// _mCpta.Open();
/// ```
fn demo_csharp_connection_style() -> SageResult<()> {
    println!("\n🎯 === REPRODUCTION EXACTE DU CODE C# SAGE ===");
    
    // Équivalent C# : private BSCPTAApplication100c _mCpta;
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    println!("✅ BSCPTAApplication100c créée");
    
    // Équivalent C# : _mCpta.Name = "D:\\TMP\\BIJOU.MAE";
    app.set_name(r"D:\TMP\BIJOU.MAE")?;
    println!("✅ Base définie: D:\\TMP\\BIJOU.MAE");
    
    // Vérification que le nom a été défini
    let current_name = app.name()?;
    println!("📋 Nom actuel: '{}'", current_name);
    
    // Équivalent C# : _mCpta.Loggable.UserName = "<Administrateur>";
    let loggable = app.loggable()?;
    loggable.set_user_name("<Administrateur>")?;
    println!("✅ Nom d'utilisateur défini: <Administrateur>");
    
    // Équivalent C# : _mCpta.Loggable.UserPwd = "";
    loggable.set_user_pwd("")?;
    println!("✅ Mot de passe défini: (vide)");
    
    // Vérification que les credentials ont été définis
    let current_username = loggable.user_name()?;
    println!("👤 Utilisateur actuel: '{}'", current_username);
    
    // Équivalent C# : _mCpta.Open();
    println!("\n🔑 Tentative de connexion...");
    match app.open() {
        Ok(()) => {
            println!("🎉 CONNEXION RÉUSSIE!");
            
            // Vérifier le statut de connexion
            if app.is_open()? {
                println!("✅ Base de données ouverte");
                
                // Afficher les informations de connexion
                println!("📊 Informations de connexion:");
                println!("   - Base: {}", app.name()?);
                println!("   - Utilisateur: {}", app.loggable()?.user_name()?);
                println!("   - Connecté: {}", app.loggable()?.is_logged()?);
                println!("   - Admin: {}", app.loggable()?.is_administrator()?);
                
                // Fermer proprement
                app.close()?;
                println!("✅ Base fermée proprement");
            } else {
                println!("⚠️ Base pas ouverte malgré le succès d'Open()");
            }
        }
        Err(e) => {
            println!("❌ ÉCHEC DE CONNEXION: {}", e);
            println!("💡 Causes possibles:");
            println!("   - Fichier BIJOU.MAE inexistant");
            println!("   - Credentials incorrects");
            println!("   - Base corrompue");
            println!("   - Sage 100c non installé");
        }
    }
    
    println!("\n🎯 === COMPARAISON C# vs RUST ===");
    println!("C# Style  : _mCpta.Name = \"D:\\\\TMP\\\\BIJOU.MAE\";");
    println!("Rust Style: app.set_name(r\"D:\\TMP\\BIJOU.MAE\")?;");
    println!();
    println!("C# Style  : _mCpta.Loggable.UserName = \"<Administrateur>\";");
    println!("Rust Style: app.loggable()?.set_user_name(\"<Administrateur>\")?;");
    println!();
    println!("C# Style  : _mCpta.Open();");
    println!("Rust Style: app.open()?;");
    
    Ok(())
}

/// Fonction utilitaire pour tester différents chemins de base
#[allow(dead_code)]
fn test_different_database_paths() -> SageResult<()> {
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    
    let test_paths = vec![
        r"D:\TMP\BIJOU.MAE",
        r"C:\Sage\DEMO.MAE", 
        r".\DEMO.MAE",
    ];
    
    for path in test_paths {
        println!("🧪 Test du chemin: {}", path);
        
        match app.set_name(path) {
            Ok(()) => println!("  ✅ Chemin défini"),
            Err(e) => println!("  ❌ Erreur: {}", e),
        }
        
        match app.name() {
            Ok(name) => println!("  📋 Nom lu: '{}'", name),
            Err(e) => println!("  ❌ Lecture échouée: {}", e),
        }
        
        println!();
    }
    
    Ok(())
}

/// Fonction utilitaire pour tester différents credentials
#[allow(dead_code)] 
fn test_different_credentials() -> SageResult<()> {
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    let loggable = app.loggable()?;
    
    let test_credentials = vec![
        ("<Administrateur>", ""),
        ("admin", "password"),
        ("", ""),
    ];
    
    for (username, password) in test_credentials {
        println!("🧪 Test credentials: '{}' / '{}'", username, password);
        
        match loggable.set_user_name(username) {
            Ok(()) => println!("  ✅ Username défini"),
            Err(e) => println!("  ❌ Erreur username: {}", e),
        }
        
        match loggable.set_user_pwd(password) {
            Ok(()) => println!("  ✅ Password défini"),
            Err(e) => println!("  ❌ Erreur password: {}", e),
        }
        
        println!();
    }
    
    Ok(())
}
