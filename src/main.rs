mod com;
mod errors;

use com::{ComInstance, SafeDispatch, SafeVariant, MemberType};
#[allow(unused_imports)] // Sera utilisé dans les futures versions  
use com::MemberInfo;
use errors::SageResult;

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

#[allow(dead_code)] // A supprimer à la finalisation de la v0.2.0
const BSCIAL_CLSID: &str = "ED0EC116-16B8-44CC-A68A-41BF6E15EB3F";

fn main() -> SageResult<()> {
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

    Ok(())
}

fn format_variant_result(variant: &SafeVariant) -> String {
    match variant.to_string() {
        Ok(s) => format!("{} ({})", s, variant.type_name()),
        Err(_) => format!("Valeur {} non convertible en string", variant.type_name()),
    }
}
