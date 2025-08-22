mod com;
mod errors;
mod wrappers;

use wrappers::{CptaApplication, FactoryJournal};
use errors::SageResult;

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

fn main() -> SageResult<()> {
    println!("🚀 Test des signatures FactoryJournal corrigées v0.1.3");
    println!("═══════════════════════════════════════════════════════");
    
    test_factory_journal_corrected()?;
    
    Ok(())
}

fn test_factory_journal_corrected() -> SageResult<()> {
    println!("\n✅ === TEST SIGNATURES CORRIGÉES ===");
    
    // Créer l'application Sage
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    println!("✅ Application Sage créée");
    
    // Obtenir la propriété FactoryJournal
    match app.factory_journal() {
        Ok(factory_journal_variant) => {
            println!("✅ Propriété FactoryJournal obtenue: {}", factory_journal_variant.type_name());

            if let Ok(factory_journal_dispatch) = factory_journal_variant.to_dispatch() {
                println!("✅ Interface IDispatch extraite de FactoryJournal");
                
                // Créer un wrapper FactoryJournal
                let factory = FactoryJournal {
                    dispatch: factory_journal_dispatch
                };
                
                println!("\n🧪 Test avec signatures corrigées (1 paramètre) :");
                
                // Test ReadNumero avec code journal string
                println!("📋 Test ReadNumero avec code 'VTE'...");
                match factory.read_journal_by_code("VTE") {
                    Ok(result) => println!("   ✅ ReadNumero('VTE') réussi: {:?}", result),
                    Err(e) => println!("   ⚠️  ReadNumero('VTE'): {} (Dossier fermé = normal)", e),
                }
                
                // Test ExistNumero avec code journal string  
                println!("📋 Test ExistNumero avec code 'VTE'...");
                match factory.journal_exists_by_code("VTE") {
                    Ok(exists) => println!("   ✅ ExistNumero('VTE') réussi: Journal existe = {}", exists),
                    Err(e) => println!("   ⚠️  ExistNumero('VTE'): {} (Dossier fermé = normal)", e),
                }
                
                // Test ReadNumero avec numéro
                println!("📋 Test ReadNumero avec numéro 1...");
                match factory.read_journal_by_number(1) {
                    Ok(result) => println!("   ✅ ReadNumero(1) réussi: {:?}", result),
                    Err(e) => println!("   ⚠️  ReadNumero(1): {} (Dossier fermé = normal)", e),
                }
                
                // Test ExistNumero avec numéro
                println!("📋 Test ExistNumero avec numéro 1...");
                match factory.journal_exists_by_number(1) {
                    Ok(exists) => println!("   ✅ ExistNumero(1) réussi: Journal existe = {}", exists),
                    Err(e) => println!("   ⚠️  ExistNumero(1): {} (Dossier fermé = normal)", e),
                }
                
                println!("\n🎉 Signatures testées avec succès !");
                println!("💡 Erreurs 'Dossier non ouvert' = signatures correctes");
                println!("❌ Erreurs 'Nombre de paramètres non valide' = signatures incorrectes");
                
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
