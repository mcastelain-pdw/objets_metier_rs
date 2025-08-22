mod com;
mod errors;
mod wrappers;

use com::{ComInstance, SafeDispatch, SafeVariant};
use wrappers::CptaApplication;
use errors::SageResult;

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

fn main() -> SageResult<()> {
    println!("🚀 Sage 100c - Interface Rust v0.1.3 ✅");
    println!("═══════════════════════════════════════════");
    
    test_factory_journal_signatures()?;
    
    Ok(())
}

fn test_factory_journal_signatures() -> SageResult<()> {
    println!("\n🔍 === ANALYSE DES SIGNATURES FACTORY JOURNAL ===");
    
    // Créer l'application Sage
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    println!("✅ Application Sage créée");
    
    // Obtenir le SafeDispatch de l'application
    let app_dispatch = app.dispatch();
    
    // Obtenir la propriété FactoryJournal
    match app_dispatch.call_method_by_name("FactoryJournal", &[]) {
        Ok(factory_journal_variant) => {
            println!("✅ Propriété FactoryJournal obtenue: {}", factory_journal_variant.type_name());

            if let Ok(factory_journal_dispatch) = factory_journal_variant.to_dispatch() {
                println!("✅ Interface IDispatch extraite de FactoryJournal");
                
                // Tester ReadNumero et ExistNumero avec différents nombres de paramètres
                test_method_signatures(&factory_journal_dispatch)?;
                
            } else {
                println!("❌ Impossible d'extraire IDispatch de la propriété FactoryJournal");
            }
        }
        Err(e) => {
            println!("❌ Impossible d'obtenir la propriété FactoryJournal: {}", e);
        }
    }
    
    Ok(())
}

fn test_method_signatures(dispatch: &windows::Win32::System::Com::IDispatch) -> SageResult<()> {
    let safe_dispatch = SafeDispatch::new(dispatch);
    
    let methods = ["ReadNumero", "ExistNumero"];
    
    for method_name in &methods {
        println!("\n🎯 === TEST MÉTHODE: {} ===", method_name);
        
        // Test avec 0 paramètres
        println!("🧪 Test avec 0 paramètres...");
        match safe_dispatch.call_method_by_name(method_name, &[]) {
            Ok(result) => println!("   ✅ Succès: {}", format_variant(&result)),
            Err(e) => println!("   ❌ Échec: {}", e),
        }
        
        // Test avec 1 paramètre string
        println!("🧪 Test avec 1 paramètre String...");
        let param1 = SafeVariant::from_string("VTE");
        match safe_dispatch.call_method_by_name(method_name, &[param1]) {
            Ok(result) => println!("   ✅ Succès: {}", format_variant(&result)),
            Err(e) => println!("   ❌ Échec: {}", e),
        }
        
        // Test avec 1 paramètre numérique
        println!("🧪 Test avec 1 paramètre I4...");
        let param2 = SafeVariant::I4(1);
        match safe_dispatch.call_method_by_name(method_name, &[param2]) {
            Ok(result) => println!("   ✅ Succès: {}", format_variant(&result)),
            Err(e) => println!("   ❌ Échec: {}", e),
        }
        
        // Test avec 2 paramètres
        println!("🧪 Test avec 2 paramètres...");
        let param1 = SafeVariant::from_string("VTE");
        let param2 = SafeVariant::I4(1);
        match safe_dispatch.call_method_by_name(method_name, &[param1, param2]) {
            Ok(result) => println!("   ✅ Succès: {}", format_variant(&result)),
            Err(e) => println!("   ❌ Échec: {}", e),
        }
    }
    
    Ok(())
}

fn format_variant(variant: &SafeVariant) -> String {
    match variant {
        SafeVariant::Empty => "Empty".to_string(),
        SafeVariant::Null => "Null".to_string(),
        SafeVariant::I4(i) => format!("I4({})", i),
        SafeVariant::BStr(s) => format!("String(\"{}\")", s),
        SafeVariant::Bool(b) => format!("Bool({})", b),
        SafeVariant::Dispatch(_) => "Object(IDispatch)".to_string(),
        SafeVariant::Error(e) => format!("Error({})", e),
        other => format!("{:?}", other),
    }
}
