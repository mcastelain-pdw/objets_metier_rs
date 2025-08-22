mod com;
mod errors;
mod wrappers;

use com::{ComInstance, SafeDispatch, SafeVariant, MemberType};
use wrappers::CptaApplication;
#[allow(unused_imports)] // Sera utilisé dans les futures versions  
use com::MemberInfo;
use errors::SageResult;

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

#[allow(dead_code)] // A supprimer à la finalisation de la v0.2.0
const BSCIAL_CLSID: &str = "ED0EC116-16B8-44CC-A68A-41BF6E15EB3F";

fn main() -> SageResult<()> {
    println!("🚀 Sage 100c - Interface Rust v0.1.3 ✅");
    println!("═══════════════════════════════════════════");
    println!("🎉 Architecture modulaire + Conversion VARIANT complète");
    println!("📦 Wrappers: CptaApplication, CptaLoggable");
    println!("🔗 Connexion Sage fonctionnelle validée");
    
    // DÉMO 1: Nouvelle syntaxe élégante v0.1.3 (recommandée)
    println!("\n✨ === SYNTAXE ÉLÉGANTE v0.1.3 (Style C#) ===");
    elegant_rust_demo()?;
    
    // DÉMO 2: Ancienne approche COM directe (pour comparaison technique)
    println!("\n🔧 === APPROCHE COM DIRECTE (Comparaison technique) ===");
    classic_com_demo()?;
    
    println!("\n🎯 === PROCHAINES ÉTAPES ===");
    println!("📋 v0.2.0: Module Commercial (CialApplication)");
    println!("💰 v0.3.0: Module Paie (PaieApplication)");
    println!("🏭 v1.0.0: Production Ready avec tous modules");
    
    Ok(())
}

/// NOUVELLE APPROCHE v0.1.3: Syntaxe élégante style C#
/// 
/// Reproduit l'expérience développeur C# Sage:
/// ```csharp
/// var app = new BSCPTAApplication100c();
/// app.Name = @"D:\TMP\BIJOU.MAE";
/// app.Loggable.UserName = "<Administrateur>";
/// app.Open();
/// ```
/// 
/// Équivalent Rust v0.1.3:
/// ```rust
/// let app = CptaApplication::new(BSCPTA_CLSID)?;
/// app.set_name(r"D:\TMP\BIJOU.MAE")?;
/// app.loggable()?.set_user_name("<Administrateur>")?;
/// app.open()?;
/// ```
/// Équivalent C#/VB : app.Loggable.UserName
/// Syntaxe Rust     : app.loggable()?.user_name()?
fn elegant_rust_demo() -> SageResult<()> {
    // Création avec la nouvelle API élégante
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    println!("✅ Application Sage créée avec CptaApplication");
    
    // Propriétés de base - syntaxe simple
    println!("📋 Nom: '{}'", app.name()?);
    println!("🔓 Base ouverte: {}", app.is_open()?);
    
    // 🎯 MAGIE! Accès aux sous-objets COM style C#/VB
    println!("\n🎯 Accès aux propriétés Loggable:");
    
    // Équivalent C#/VB: app.Loggable.UserName
    let username = app.loggable()?.user_name()?;
    println!("👤 app.loggable()?.user_name()? = '{}'", username);
    
    // Équivalent C#/VB: app.Loggable.IsLogged  
    let is_logged = app.loggable()?.is_logged()?;
    println!("🔐 app.loggable()?.is_logged()? = {}", is_logged);
    
    // Équivalent C#/VB: app.Loggable.IsAdministrator
    let is_admin = app.loggable()?.is_administrator()?;
    println!("� app.loggable()?.is_administrator()? = {}", is_admin);
    
    // Méthode helper qui combine plusieurs propriétés
    println!("📊 {}", app.loggable()?.user_info()?);
    
    // CHAÎNAGE DIRECT possible! Mais on évite ServiceName qui peut ne pas exister
    // let service = app.loggable()?.service_name()?;
    // println!("🔧 app.loggable()?.service_name()? = '{}'", service);
    
    // À la place, testons une propriété qui existe
    let user_info = app.loggable()?.user_info()?;
    println!("📋 Résumé utilisateur: {}", user_info);
    
    println!("\n🎉 Syntaxe réussie! Rust peut faire du C#/VB style!");
    
    Ok(())
}

/// ANCIENNE APPROCHE: COM direct (pour comparaison)
fn classic_com_demo() -> SageResult<()> {
    // Créer l'instance COM avec gestion automatique
    let instance = ComInstance::new(BSCPTA_CLSID)?;
    println!("✅ Instance BSCPTAApplication100c créée avec succès !");

    // Obtenir les informations de type
    match instance.get_type_info() {
        Ok(info) => println!("📋 {}", info),
        Err(e) => println!("⚠️  Impossible d'obtenir les infos de type: {}", e),
    }

    // Lister séparément méthodes et propriétés
    display_methods_and_properties(&instance)?;

    // Tester les appels de méthodes sûrs
    if instance.supports_automation() {
        println!("\n🔍 Test des appels de méthodes sûrs...");
        test_safe_method_calls(&instance)?;
    }

    println!("✅ Instance libérée automatiquement");
    Ok(())
}

fn display_methods_and_properties(instance: &ComInstance) -> SageResult<()> {
    // Afficher les méthodes
    match instance.list_methods_only() {
        Ok(methods) => {
            println!("\n🔧 MÉTHODES disponibles ({} trouvées):", methods.len());
            for method in methods.iter() {
                let params = method.param_count.map_or_else(
                    || "?".to_string(),
                    |count| count.to_string()
                );
                let return_type = method.return_type.as_deref().unwrap_or("?");
                println!("   [{}] {}({} params) -> {}", 
                    method.id, method.name, params, return_type);
            }
        }
        Err(e) => println!("⚠️  Impossible de lister les méthodes: {}", e),
    }

    // Afficher les propriétés groupées
    match instance.group_properties() {
        Ok(properties) => {
            println!("\n📋 PROPRIÉTÉS disponibles ({} trouvées):", properties.len());
            for (name, variants) in properties.iter() {
                let types: Vec<String> = variants.iter().map(|v| {
                    match v.member_type {
                        MemberType::PropertyGet => "get".to_string(),
                        MemberType::PropertyPut => "put".to_string(),
                        MemberType::PropertyPutRef => "putref".to_string(),
                        _ => "?".to_string(),
                    }
                }).collect();
                
                let return_type = variants.first()
                    .and_then(|v| v.return_type.as_deref())
                    .unwrap_or("?");
                    
                let id = variants.first().map_or(0, |v| v.id);
                
                println!("   [{}] {} [{}] -> {}", 
                    id, name, types.join("/"), return_type);
            }
        }
        Err(e) => println!("⚠️  Impossible de lister les propriétés: {}", e),
    }

    Ok(())
}

fn test_safe_method_calls(instance: &ComInstance) -> SageResult<()> {
    let dispatch = instance.dispatch()?;
    let safe_dispatch = SafeDispatch::new(dispatch);

    // Tester quelques propriétés communes
    let test_properties = [
        ("IsOpen", "Vérifier si une base est ouverte"),
        ("Name", "Nom de l'application"),
        ("Version", "Version de l'application"),
    ];

    for (prop_name, description) in test_properties {
        match safe_dispatch.call_method_by_name(prop_name, &[]) {
            Ok(result) => {
                println!("✅ {} ({}): {}", 
                    prop_name, description, format_variant_result(&result));
            }
            Err(e) => {
                println!("❌ {} ({}): {}", prop_name, description, e);
            }
        }
    }

    // Test spécial pour la propriété Loggable (IBILoggable)
    println!("\n🔍 Test de la propriété Loggable (IBILoggable)...");

    test_loggable_property(&safe_dispatch)?;

    Ok(())
}

fn test_loggable_property(safe_dispatch: &SafeDispatch) -> SageResult<()> {
    match safe_dispatch.call_method_by_name("Loggable", &[]) {
        Ok(loggable_variant) => {
            println!("✅ Propriété Loggable obtenue: {}", loggable_variant.type_name());
            
            // Vérifier si c'est un objet COM
            if loggable_variant.is_object() {
                println!("✅ Loggable est bien un objet COM");
                
                // Essayer d'extraire l'interface IDispatch
                if let Ok(loggable_dispatch) = loggable_variant.to_dispatch() {
                    println!("✅ Interface IDispatch extraite de Loggable");
                    
                    // Explorer automatiquement l'objet
                    ComInstance::explore_nested_object(loggable_dispatch.clone())?;
                    
                    // Créer un SafeDispatch pour l'objet IBILoggable
                    let loggable_safe = SafeDispatch::new(&loggable_dispatch);
                    
                    // Tester les 4 propriétés de IBILoggable
                    println!("\n🔍 Test des propriétés IBILoggable:");
                    let loggable_properties = [
                        ("IsAdministrator", "Indique si l'utilisateur est administrateur"),
                        ("IsLogged", "Indique si un utilisateur est connecté"),
                        ("UserName", "Nom d'utilisateur connecté"),
                        ("UserPwd", "Mot de passe utilisateur"),
                    ];
                    
                    for (prop_name, description) in loggable_properties {
                        match loggable_safe.call_method_by_name(prop_name, &[]) {
                            Ok(result) => {
                                println!("   ✅ {} ({}): {}", 
                                    prop_name, description, format_variant_result(&result));
                            }
                            Err(e) => {
                                println!("   ❌ {} ({}): {}", prop_name, description, e);
                            }
                        }
                    }
                } else {
                    println!("❌ Impossible d'extraire IDispatch de la propriété Loggable");
                }
            } else {
                println!("❌ Loggable n'est pas un objet COM: {}", loggable_variant.type_name());
            }
        }
        Err(e) => {
            println!("❌ Impossible d'obtenir la propriété Loggable: {}", e);
        }
    }
    
    Ok(())
}

fn format_variant_result(variant: &SafeVariant) -> String {
    match variant.to_string() {
        Ok(s) => format!("{} ({})", s, variant.type_name()),
        Err(_) => format!("Valeur {} non convertible en string", variant.type_name()),
    }
}
