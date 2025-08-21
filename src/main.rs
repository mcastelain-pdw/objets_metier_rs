mod com;
mod errors;

use com::{ComInstance, SafeDispatch, SafeVariant};
use errors::SageResult;

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

fn main() -> SageResult<()> {
    // Créer l'instance COM avec gestion automatique
    let instance = ComInstance::new(BSCPTA_CLSID)?;
    println!("✅ Instance BSCPTAApplication100c créée avec succès !");

    // Obtenir les informations de type
    match instance.get_type_info() {
        Ok(info) => println!("📋 {}", info),
        Err(e) => println!("⚠️  Impossible d'obtenir les infos de type: {}", e),
    }

    // Lister les méthodes disponibles
    match instance.list_methods() {
        Ok(methods) => {
            println!("🔧 Méthodes disponibles:");
            for (id, name) in methods.iter().take(10) {
                // Limiter à 10 pour l'exemple
                println!("   [{}] {}", id, name);
            }
        }
        Err(e) => println!("⚠️  Impossible de lister les méthodes: {}", e),
    }

    // Tester les appels de méthodes sûrs
    if instance.supports_automation() {
        println!("🔍 Test des appels de méthodes sûrs...");
        test_safe_method_calls(&instance)?;
    }

    println!("✅ Instance libérée automatiquement");
    Ok(())
}

fn test_safe_method_calls(instance: &ComInstance) -> SageResult<()> {
    let dispatch = instance.dispatch()?;
    let safe_dispatch = SafeDispatch::new(dispatch);

    // Test des méthodes communes
    let test_methods = [(1, "IsOpen"), (10, "Name"), (5, "DatabaseInfo")];

    for (method_id, method_name) in test_methods {
        match safe_dispatch.call_method(method_id, method_name) {
            Ok(result) => {
                println!("✅ {}(): {}", method_name, format_variant_result(&result));
            }
            Err(e) => {
                println!("❌ {}(): {}", method_name, e);
            }
        }
    }

    // Test d'appel par nom si possible
    match safe_dispatch.call_method_by_name("IsOpen", &[]) {
        Ok(result) => {
            println!("✅ IsOpen() par nom: {}", format_variant_result(&result));
        }
        Err(e) => {
            println!("ℹ️  Appel par nom non supporté: {}", e);
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
