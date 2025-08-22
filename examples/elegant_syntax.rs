use objets_metier_rs::wrappers::CptaApplication;
use objets_metier_rs::errors::SageResult;

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

#[allow(dead_code)] // A supprimer à la finalisation de la v0.2.0
const BSCIAL_CLSID: &str = "ED0EC116-16B8-44CC-A68A-41BF6E15EB3F";

fn main() -> SageResult<()> {
    println!("🚀 Sage 100c - Interface Rust Élégante v0.1.3");
    
    // APPROCHE 1: Syntaxe Rust élégante similaire à C#/VB
    elegant_rust_approach()?;
    
    // APPROCHE 2: Comparaison avec l'ancienne approche
    println!("\n{}", "=".repeat(50).as_str());
    println!("🔄 Comparaison avec l'ancienne approche COM directe...");
    
    Ok(())
}

/// NOUVELLE APPROCHE : Syntaxe Rust élégante
/// Équivalent C#/VB : MaDLL.Loggable.UserName
/// Syntaxe Rust     : app.loggable()?.user_name()?
fn elegant_rust_approach() -> SageResult<()> {
    println!("\n✨ === NOUVELLE SYNTAXE RUST ÉLÉGANTE ===");
    
    // Création de l'application - Équivalent : var app = new BSCPTAApplication()
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    println!("✅ Application Sage créée");
    
    // Accès direct aux propriétés de base - Équivalent : app.Name, app.IsOpen
    println!("📋 Nom de l'application: '{}'", app.get_name()?);
    println!("🔓 Base ouverte: {}", app.is_open()?);
    
    // SYNTAX MAGIQUE! Équivalent C#/VB : app.Loggable.UserName
    // En Rust : app.loggable()?.user_name()?
    println!("\n🎯 === ACCÈS AUX SOUS-OBJETS COM ===");
    
    let loggable = app.loggable()?;
    println!("✅ Objet Loggable obtenu");
    
    // Équivalent C#/VB : app.Loggable.UserName
    let username = loggable.get_user_name()?;
    println!("👤 Nom d'utilisateur: '{}'", username);
    
    // Équivalent C#/VB : app.Loggable.IsLogged
    let is_logged = loggable.is_logged()?;
    println!("🔐 Utilisateur connecté: {}", is_logged);
    
    // Équivalent C#/VB : app.Loggable.IsAdministrator
    let is_admin = loggable.is_administrator()?;
    println!("👑 Administrateur: {}", is_admin);
    
    // Méthode helper pour afficher toutes les infos utilisateur
    println!("📊 Résumé utilisateur: {}", loggable.user_info()?);
    
    // CHAÎNAGE POSSIBLE! Équivalent C#/VB : app.Loggable.ServiceName
    let service_name = app.loggable()?.service_name()?;
    println!("🔧 Service: '{}'", service_name);
    
    // SYNTAXE ULTRA-CONDENSÉE! Une ligne équivalente à C#/VB
    let user_summary = format!(
        "Utilisateur '{}' {} (Admin: {})",
        app.loggable()?.get_user_name()?,
        if app.loggable()?.is_logged()? { "connecté" } else { "déconnecté" },
        app.loggable()?.is_administrator()?
    );
    println!("📝 Résumé: {}", user_summary);
    
    println!("\n🎉 === SYNTAXE RÉUSSIE! ===");
    println!("✅ C#/VB  : app.Loggable.UserName");
    println!("✅ Rust   : app.loggable()?.user_name()?");
    println!("✅ Chaîne : app.loggable()?.service_name()?");
    
    Ok(())
}

/// Fonction helper pour montrer les différentes syntaxes possibles
#[allow(dead_code)]
fn syntax_examples() -> SageResult<()> {
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    
    // === EXEMPLES DE SYNTAXES POSSIBLES ===
    
    // Style fonctionnel avec gestion d'erreur
    let username = app.loggable()
        .and_then(|l| l.get_user_name())
        .unwrap_or_else(|_| "Utilisateur inconnu".to_string());
    
    // Style impératif
    let loggable = app.loggable()?;
    let is_admin = loggable.is_administrator()?;
    let is_logged = loggable.is_logged()?;
    
    // Chaînage direct
    let service = app.loggable()?.service_name()?;
    
    // Pattern matching pour gestion d'erreur fine
    match app.loggable() {
        Ok(loggable) => {
            match loggable.get_user_name() {
                Ok(name) => println!("Utilisateur: {}", name),
                Err(e) => println!("Erreur nom utilisateur: {}", e),
            }
        },
        Err(e) => println!("Erreur accès Loggable: {}", e),
    }
    
    println!("Exemples: user={}, admin={}, logged={}, service={}", 
        username, is_admin, is_logged, service);
    
    Ok(())
}

/// Démonstrateur pour d'autres objets Factory
#[allow(dead_code)]
fn factory_objects_example() -> SageResult<()> {
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    
    // Dans le futur, on pourrait avoir :
    // let clients = app.factory_client()?.list()?;
    // let tiers = app.factory_tiers()?.find_by_code("CLI001")?;
    // let compte = app.factory_compte_g()?.create_new()?;
    
    println!("🚧 Fonctionnalités futures pour les Factory objects");
    
    Ok(())
}
