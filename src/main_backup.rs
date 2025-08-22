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
    println!("📋 Nom: '{}'", app.get_name()?);
    println!("🔓 Base ouverte: {}", app.is_open()?);
    
    // 🎯 MAGIE! Accès aux sous-objets COM style C#/VB
    println!("\n🎯 Accès aux propriétés Loggable:");
    
    // Équivalent C#/VB: app.Loggable.UserName
    let username = app.loggable()?.get_user_name()?;
    println!("👤 app.loggable()?.get_user_name()? = '{}'", username);

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

    test_loggable(&safe_dispatch)?;
    test_factory_journal(&safe_dispatch)?;

    Ok(())
}

fn test_loggable(safe_dispatch: &SafeDispatch) -> SageResult<()> {
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

fn test_factory_journal(safe_dispatch: &SafeDispatch) -> SageResult<()> {
    match safe_dispatch.call_method_by_name("FactoryJournal", &[]) {
        Ok(factory_journal_variant) => {
            println!("✅ Propriété FactoryJournal obtenue: {}", factory_journal_variant.type_name());

            // Vérifier si c'est un objet COM
            if factory_journal_variant.is_object() {
                println!("✅ FactoryJournal est bien un objet COM");

                // Essayer d'extraire l'interface IDispatch
                if let Ok(factory_journal_dispatch) = factory_journal_variant.to_dispatch() {
                    println!("✅ Interface IDispatch extraite de FactoryJournal");
                    
                    // Explorer automatiquement l'objet
                    ComInstance::explore_nested_object(factory_journal_dispatch.clone())?;

                    // NOUVELLE ANALYSE DÉTAILLÉE DES MÉTHODES
                    println!("\n🔍 === ANALYSE DÉTAILLÉE DES MÉTHODES IBOJournalFactory3 ===");
                    analyze_factory_journal_methods(&factory_journal_dispatch)?;

                } else {
                    println!("❌ Impossible d'extraire IDispatch de la propriété FactoryJournal");
                }
            } else {
                println!("❌ FactoryJournal n'est pas un objet COM: {}", factory_journal_variant.type_name());
            }
        }
        Err(e) => {
            println!("❌ Impossible d'obtenir la propriété FactoryJournal: {}", e);
        }
    }
    
    Ok(())
}

fn analyze_factory_journal_methods(dispatch: &windows::Win32::System::Com::IDispatch) -> SageResult<()> {
    use windows::Win32::System::Com::{ITypeInfo, TYPEATTR, FUNCDESC};
    
    unsafe {
        // Obtenir l'interface ITypeInfo pour les signatures détaillées
        if let Ok(type_info) = dispatch.GetTypeInfo(0, 0) {
            println!("📋 Analyse des signatures de méthodes...");
            
            // Obtenir les attributs du type
            let mut type_attr_ptr = std::ptr::null_mut();
            if type_info.GetTypeAttr(&mut type_attr_ptr).is_ok() {
                let type_attr = &*type_attr_ptr;
                
                println!("📊 Type: {:?}, {} fonctions trouvées", type_attr.typekind, type_attr.cFuncs);
                
                // Analyser chaque fonction
                for i in 0..type_attr.cFuncs {
                    let mut func_desc_ptr = std::ptr::null_mut();
                    if type_info.GetFuncDesc(i, &mut func_desc_ptr).is_ok() {
                        let func_desc = &*func_desc_ptr;
                        
                        // Obtenir le nom de la fonction
                        let mut names = vec![windows::core::BSTR::default(); 1];
                        let mut names_count = 0u32;
                        
                        if type_info.GetNames(func_desc.memid, names.as_mut_ptr(), 1, &mut names_count).is_ok() && names_count > 0 {
                            let method_name = names[0].to_string();
                            
                            // Analyser spécifiquement ReadNumero et ExistNumero
                            if method_name == "ReadNumero" || method_name == "ExistNumero" {
                                println!("\n🎯 === MÉTHODE: {} ===", method_name);
                                println!("📋 DISPID: {}", func_desc.memid);
                                println!("📋 Invoke Kind: {:?}", func_desc.invkind);
                                println!("📋 Nombre de paramètres: {}", func_desc.cParams);
                                println!("📋 Paramètres optionnels: {}", func_desc.cParamsOpt);
                                
                                // Type de retour
                                let return_type = format_var_type(&func_desc.elemdescFunc.tdesc);
                                println!("📋 Type de retour: {}", return_type);
                                
                                // Paramètres
                                if func_desc.cParams > 0 {
                                    println!("📋 Paramètres:");
                                    let params_slice = std::slice::from_raw_parts(
                                        func_desc.lprgelemdescParam,
                                        func_desc.cParams as usize
                                    );
                                    
                                    for (param_idx, param) in params_slice.iter().enumerate() {
                                        let param_type = format_var_type(&param.tdesc);
                                        println!("   [{}] Type: {}", param_idx, param_type);
                                    }
                                    
                                    // Essayer d'obtenir les noms des paramètres
                                    let mut param_names = vec![windows::core::BSTR::default(); (func_desc.cParams + 1) as usize];
                                    let mut param_names_count = 0u32;
                                    
                                    if type_info.GetNames(
                                        func_desc.memid, 
                                        param_names.as_mut_ptr(), 
                                        (func_desc.cParams + 1) as u32, 
                                        &mut param_names_count
                                    ).is_ok() && param_names_count > 1 {
                                        println!("📋 Noms des paramètres:");
                                        for i in 1..param_names_count as usize {
                                            if i <= func_desc.cParams as usize {
                                                println!("   [{}] {}", i-1, param_names[i].to_string());
                                            }
                                        }
                                    }
                                }
                                
                                // Test d'appel avec le bon nombre de paramètres
                                test_method_with_params(dispatch, &method_name, func_desc.cParams as usize)?;
                            }
                        }
                        
                        // Libérer la description de fonction
                        type_info.ReleaseFuncDesc(func_desc_ptr);
                    }
                }
                
                // Libérer les attributs du type
                type_info.ReleaseTypeAttr(type_attr_ptr);
            }
        } else {
            println!("❌ Impossible d'obtenir ITypeInfo pour l'analyse détaillée");
        }
    }
    
    Ok(())
}

fn format_var_type(type_desc: &windows::Win32::System::Com::TYPEDESC) -> String {
    use windows::Win32::System::Variant::{
        VT_BSTR, VT_I4, VT_BOOL, VT_DISPATCH, VT_VARIANT, VT_VOID, VT_HRESULT
    };
    
    unsafe {
        match type_desc.Anonymous.vt {
            VT_VOID => "void".to_string(),
            VT_I4 => "long".to_string(),
            VT_BSTR => "BSTR".to_string(),
            VT_BOOL => "VARIANT_BOOL".to_string(),
            VT_DISPATCH => "IDispatch*".to_string(),
            VT_VARIANT => "VARIANT".to_string(),
            VT_HRESULT => "HRESULT".to_string(),
            other => format!("VT_{}", other.0),
        }
    }
}

fn test_method_with_params(dispatch: &windows::Win32::System::Com::IDispatch, method_name: &str, param_count: usize) -> SageResult<()> {
    let safe_dispatch = SafeDispatch::new(dispatch);
    
    println!("🧪 Test d'appel de {} avec {} paramètres...", method_name, param_count);
    
    match param_count {
        0 => {
            // Aucun paramètre
            match safe_dispatch.call_method_by_name(method_name, &[]) {
                Ok(result) => println!("   ✅ Succès: {}", format_variant_result(&result)),
                Err(e) => println!("   ❌ Échec: {}", e),
            }
        }
        1 => {
            // 1 paramètre - essayer différents types courants
            let test_params = vec![
                SafeVariant::from_string("VTE"),  // Code journal typique
                SafeVariant::I4(1),              // Numéro
                SafeVariant::Empty,              // Vide
            ];
            
            for param in test_params {
                println!("   🧪 Essai avec paramètre: {:?}", param);
                match safe_dispatch.call_method_by_name(method_name, &[param]) {
                    Ok(result) => {
                        println!("   ✅ Succès: {}", format_variant_result(&result));
                        break; // Arrêter au premier succès
                    }
                    Err(e) => println!("   ❌ Échec: {}", e),
                }
            }
        }
        _ => {
            println!("   ⚠️  Méthode avec {} paramètres - test manuel requis", param_count);
        }
    }
    
    Ok(())
}

fn format_variant_result(variant: &SafeVariant) -> String {
    match variant {
        SafeVariant::Empty => "Empty".to_string(),
        SafeVariant::Null => "Null".to_string(),
        SafeVariant::I4(i) => format!("I4({})", i),
        SafeVariant::BStr(s) => format!("String("{}")", s),
        SafeVariant::Bool(b) => format!("Bool({})", b),
        SafeVariant::Dispatch(_) => "Object(IDispatch)".to_string(),
        SafeVariant::Error(e) => format!("Error({})", e),
        other => format!("{:?}", other),
    }
}

fn format_variant_result(variant: &SafeVariant) -> String {
    match variant {
        SafeVariant::Empty => "Empty".to_string(),
        SafeVariant::Null => "Null".to_string(),
        SafeVariant::I4(i) => format!("I4({})", i),
        SafeVariant::String(s) => format!("String(\"{}\")", s),
        SafeVariant::Bool(b) => format!("Bool({})", b),
        SafeVariant::Object(_) => "Object(IDispatch)".to_string(),
        SafeVariant::Error(e) => format!("Error({})", e),
        other => format!("{:?}", other),
    }
}

fn test_method_with_params(dispatch: &IDispatch, method_name: &str, param_count: usize) -> SageResult<()> {
    let safe_dispatch = SafeDispatch::new(dispatch);
    
    println!("🧪 Test d'appel de {} avec {} paramètres...", method_name, param_count);
    
    match param_count {
        0 => {
            // Aucun paramètre
            match safe_dispatch.call_method_by_name(method_name, &[]) {
                Ok(result) => println!("   ✅ Succès: {}", format_variant_result(&result)),
                Err(e) => println!("   ❌ Échec: {}", e),
            }
        }
        1 => {
            // 1 paramètre - essayer différents types courants
            let test_params = vec![
                SafeVariant::from_string("VTE"),  // Code journal typique
                SafeVariant::I4(1),              // Numéro
                SafeVariant::Empty,              // Vide
            ];
            
            for param in test_params {
                println!("   🧪 Essai avec paramètre: {:?}", param);
                match safe_dispatch.call_method_by_name(method_name, &[param]) {
                    Ok(result) => {
                        println!("   ✅ Succès: {}", format_variant_result(&result));
                        break; // Arrêter au premier succès
                    }
                    Err(e) => println!("   ❌ Échec: {}", e),
                }
            }
        }
        _ => {
            println!("   ⚠️  Méthode avec {} paramètres - test manuel requis", param_count);
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
