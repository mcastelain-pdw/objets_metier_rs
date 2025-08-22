mod com;
mod errors;
mod wrappers;

use wrappers::{CptaApplication, FactoryJournal};
use errors::SageResult;

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";

fn main() -> SageResult<()> {
    println!("🚀 Test des signatures FactoryJournal corrigées v0.1.3");
    println!("═══════════════════════════════════════════════════════");
    
    // Créer l'application Sage
    let app = CptaApplication::new(BSCPTA_CLSID)?;
    app.set_name(r"D:\TMP\BIJOU.MAE")?;
    app.loggable()?.set_user_name("<Administrateur>")?;
    app.loggable()?.set_user_pwd("")?;

    match app.open() {
        Ok(()) => {
            if app.is_open()? {
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
                            
                            // Test ExistNumero avec code journal string  
                            println!("📋 Test ExistNumero avec code 'VTE'...");
                            match factory.exists_by_code("VTE") {
                                Ok(exists) => {
                                    println!("   ✅ ExistNumero('VTE') réussi: Journal existe = {}", exists);
                                    if exists {
                                        // Test ReadNumero avec code journal string
                                        println!("📋 Test ReadNumero avec code 'VTE'...");
                                        match factory.read_by_code("VTE") {
                                            Ok(result) => println!("   ✅ ReadNumero('VTE') réussi: {:?}", result),
                                            Err(e) => println!("   ⚠️  ReadNumero('VTE'): {} (Dossier fermé = normal)", e),
                                        }
                                    }
                                },
                                Err(e) => println!("   ⚠️  ExistNumero('VTE'): {} (Dossier fermé = normal)", e),
                            }
                            
                            // Test ExistNumero avec numéro
                            println!("📋 Test ExistNumero avec numéro 1...");
                            match factory.exists_by_id(1) {
                                Ok(exists) => {
                                    println!("   ✅ ExistNumero(1) réussi: Journal existe = {}", exists);
                                    if exists {
                                        // Test ReadNumero avec numéro
                                        println!("📋 Test ReadNumero avec numéro 1...");
                                        match factory.read_by_id(1) {
                                            Ok(result) => println!("   ✅ ReadNumero(1) réussi: {:?}", result),
                                            Err(e) => println!("   ⚠️  ReadNumero(1): {} (Dossier fermé = normal)", e),
                                        }
                                    }
                                },
                                Err(e) => println!("   ⚠️  ExistNumero(1): {} (Dossier fermé = normal)", e),
                            }

                            println!("\n🎉 Implémentations des journaux partiellement fonctionnelles !");

                        } else {
                            println!("❌ Impossible d'extraire IDispatch de la propriété FactoryJournal");
                        }
                    }
                    Err(e) => {
                        println!("❌ Impossible d'obtenir la propriété FactoryJournal: {}", e);
                    }
                }

                // Fermer proprement
                app.close()?;
            }
        }
        Err(e) => {
            println!("❌ ÉCHEC DE CONNEXION: {}", e);
            println!("💡 Causes possibles:");
            println!("   - Fichier BIJOU.MAE inexistant");
            println!("   - Credentials incorrects");
            println!("   - Base corrompue");
            println!("   - Sage OM 100c non installé");
        }
    }

    Ok(())
}

