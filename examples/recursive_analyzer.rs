#![recursion_limit = "256"]

use std::collections::HashSet;
use std::fs::OpenOptions;
use std::io::Write;
use std::sync::Mutex;
use objets_metier_rs::com::{ComInstance, MemberType, SafeDispatch};
use objets_metier_rs::wrappers::CptaApplication;
use objets_metier_rs::errors::{self, SageResult};
use objets_metier_rs::CialApplication;
use windows::Win32::System::Com::{IDispatch, CoCreateInstance, CLSCTX_INPROC_SERVER, CLSIDFromProgID};
use windows::core::HSTRING;

// Logger global pour écrire dans le fichier
lazy_static::lazy_static! {
    static ref LOG_FILE: Mutex<std::fs::File> = {
        let file = OpenOptions::new()
            .create(true)
            .write(true)
            .truncate(true)
            .open("sage_analyzer_log.txt")
            .expect("Impossible de créer le fichier de log");
        Mutex::new(file)
    };
}

// Fonction helper pour écrire dans le log et la console
fn log_message(message: &str) {
    println!("{}", message);
    if let Ok(mut file) = LOG_FILE.lock() {
        let timestamp = chrono::Utc::now().format("%Y-%m-%d %H:%M:%S UTC");
        let _ = writeln!(file, "[{}] {}", timestamp, message);
        let _ = file.flush();
    }
}

// Macro simplifiée pour remplacer println! avec logging
macro_rules! log_println {
    () => {
        println!();
        if let Ok(mut file) = LOG_FILE.lock() {
            let _ = writeln!(file);
            let _ = file.flush();
        }
    };
    ($($arg:tt)*) => {
        log_message(&format!($($arg)*));
    };
}

const BSCPTA_CLSID: &str = "309DE0FB-9FB8-4F4E-8295-CC60C60DAA33";
const BSCIAL_CLSID: &str = "ED0EC116-16B8-44CC-A68A-41BF6E15EB3F";

// Type libraries COM Sage connus
const SAGE_TYPELIB_GUIDS: &[&str] = &[
    "{00020430-0000-0000-C000-000000000046}", // Standard OLE
    "{309DE0FB-9FB8-4F4E-8295-CC60C60DAA33}", // BSCPTAApplication100c
    "{ED0EC116-16B8-44CC-A68A-41BF6E15EB3F}", // BSCIALApplication100c
];

// ProgIDs COM Sage connus à tester
const SAGE_PROGIDS: &[&str] = &[
    "BSCPTAApplication100c",
    "BSCIALApplication100c",
];

/// Structure pour stocker les informations d'un TypeLib découvert
#[derive(Debug, Clone)]
struct TypeLibInfo {
    name: String,
    guid: String,
    version: String,
    interfaces: Vec<InterfaceInfo>,
}

/// Structure pour stocker les informations d'une interface COM
#[derive(Debug, Clone)]
struct InterfaceInfo {
    name: String,
    guid: String,
    is_dual: bool,
    methods: Vec<MethodInfo>,
    properties: Vec<PropertyInfo>,
}

/// Structure pour stocker les informations d'un objet COM analysé
#[derive(Debug, Clone)]
struct ComObjectInfo {
    name: String,
    object_type: String,
    properties: Vec<PropertyInfo>,
    methods: Vec<MethodInfo>,
    sub_objects: Vec<ComObjectInfo>,
}

#[derive(Debug, Clone)]
struct PropertyInfo {
    name: String,
    property_type: String,
    is_object: bool,
    access_type: String, // "get", "put", "get/put"
    id: i32,
    description: Option<String>, // Description enrichie
}

#[derive(Debug, Clone)]
struct MethodInfo {
    name: String,
    return_type: String,
    parameter_count: Option<u32>,
    id: i32,
    parameters: Vec<ParameterInfo>, // Nouveaux détails des paramètres
}

#[derive(Debug, Clone)]
struct ParameterInfo {
    name: Option<String>,
    param_type: String,
    is_optional: bool,
}

fn main() -> SageResult<()> {
    log_println!("🔍 Analyseur récursif des objets COM Sage 100c v0.1.3 - AVEC LOGGING");
    log_println!("═══════════════════════════════════════════════════════════════════");
    log_println!("🎯 Objectif: Cartographier TOUTE l'API Sage disponible");
    log_println!("📁 Toutes les sorties seront enregistrées dans: sage_analyzer_log.txt");
    log_println!();
    
    // PHASE 1: Scanner les Type Libraries pour découvrir tous les objets COM
    log_println!("🔎 === PHASE 1: DÉCOUVERTE DES TYPE LIBRARIES ===");
    let discovered_typelibs = scan_sage_typelibs()?;
    print_typelib_summary(&discovered_typelibs);
    
    // PHASE 2: Tester les ProgIDs connus pour instancier des objets COM directs
    log_println!("\n🧪 === PHASE 2: TEST DES PROGIDS CONNUS ===");
    let discovered_objects = test_known_progids();
    print_progid_test_results(&discovered_objects);
    
    // PHASE 3: Analyse récursive de l'application principale
    log_println!("\n🚀 === PHASE 3: ANALYSE RÉCURSIVE TRADITIONNELLE ===");
    let app_cpta = CptaApplication::new(BSCPTA_CLSID)?;
    log_println!("✅ Application Comptabilité Sage créée");
    let app_cial = CialApplication::new(BSCIAL_CLSID)?;
    log_println!("✅ Application Gestion Commerciale Sage créée");

    let app_cpta_instance = app_cpta.instance();
    let app_cial_instance = app_cial.instance();
    let app_cpta_dispatch = app_cpta_instance.dispatch()?;
    let app_cial_dispatch = app_cial_instance.dispatch()?;

    let mut visited_objects = HashSet::new();
    let root_cpta_info = analyze_com_object_recursive(
        &app_cpta_dispatch,
        "BSCPTAApplication100c".to_string(),
        0,
        &mut visited_objects
    )?;

    visited_objects = HashSet::new();
    let root_cial_info = analyze_com_object_recursive(
        &app_cial_dispatch,
        "BSCIALApplication100c".to_string(),
        0,
        &mut visited_objects
    )?;
    
    // PHASE 4: Analyse des objets retournés par les méthodes importantes
    log_println!("\n🔬 === PHASE 4: ANALYSE DES OBJETS RETOURNÉS ===");
    log_println!("\n🔬 === SAGE COMPTABILITE ===");
    app_cpta.set_name(r"D:\TMP\BIJOU.MAE")?;
    app_cpta.loggable()?.set_user_name("<Administrateur>")?;
    app_cpta.loggable()?.set_user_pwd("")?;
    match app_cpta.open() {
        Ok(()) => {
            if app_cpta.is_open()? {
                let returned_cpta_objects = analyze_cpta_returned_objects(&app_cpta_dispatch)?;
                print_returned_objects_summary(&returned_cpta_objects);

                // Fermer proprement
                app_cpta.close()?;
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

    log_println!("\n🔬 === SAGE GESTION COMMERCIALE ===");
    app_cial.set_name(r"D:\TMP\BIJOU.GCM")?;
    app_cial.loggable()?.set_user_name("<Administrateur>")?;
    app_cial.loggable()?.set_user_pwd("")?;
    match app_cial.open() {
        Ok(()) => {
            if app_cial.is_open()? {
                let returned_cial_objects = analyze_cial_returned_objects(&app_cial_dispatch)?;
                print_returned_objects_summary(&returned_cial_objects);

                // Fermer proprement
                app_cial.close()?;
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

    // PHASE 5: Synthèse et sauvegarde
    log_println!("\n📊 === PHASE 5: SYNTHÈSE FINALE ===");
    log_println!("\n🔬 === SAGE COMPTABILITE ===");
    print_analysis_summary(&root_cpta_info);

    log_println!("\n🔬 === SAGE GESTION COMMERCIALE ===");
    print_analysis_summary(&root_cial_info);

    log_println!("\n🎉 Analyse complète terminée !");
    log_println!("📁 Fichier de log généré: sage_analyzer_log.txt");
    
    Ok(())
}

/// Scanne les Type Libraries pour découvrir les interfaces COM Sage
fn scan_sage_typelibs() -> SageResult<Vec<TypeLibInfo>> {
    let mut discovered_typelibs = Vec::new();
    
    log_println!("🔍 Scanning des Type Libraries Sage...");
    
    for &guid_str in SAGE_TYPELIB_GUIDS {
        log_println!("  🧪 Test du GUID: {}", guid_str);
        
        match scan_single_typelib(guid_str) {
            Ok(Some(typelib_info)) => {
                log_println!("  ✅ TypeLib trouvée: {} (v{})", typelib_info.name, typelib_info.version);
                discovered_typelibs.push(typelib_info);
            }
            Ok(None) => {
                log_println!("  ⚠️  TypeLib vide ou non accessible");
            }
            Err(e) => {
                log_println!("  ❌ Erreur: {}", e);
            }
        }
    }
    
    Ok(discovered_typelibs)
}

/// Scanne une Type Library spécifique
fn scan_single_typelib(guid_str: &str) -> SageResult<Option<TypeLibInfo>> {
    // Note: Cette fonction est un squelette pour l'instant
    // Une implémentation complète nécessiterait d'utiliser LoadTypeLib et ITypeLib
    // Ce qui est complexe en Rust avec windows-rs
    
    log_println!("    📚 Analyse de la TypeLib {}", guid_str);
    
    // Pour l'instant, retourner None car l'implémentation complète est complexe
    // TODO: Implémenter l'analyse des TypeLib via ITypeLib
    Ok(None)
}

/// Teste les ProgIDs connus pour découvrir des objets COM directement instanciables
fn test_known_progids() -> Vec<(String, bool, Option<ComObjectInfo>)> {
    let mut results = Vec::new();
    
    log_println!("🧪 Test des ProgIDs connus...");
    
    for &progid in SAGE_PROGIDS {
        log_println!("  🔍 Test de: {}", progid);
        
        match test_single_progid(progid) {
            Ok(com_info) => {
                log_println!("  ✅ {} instancié avec succès!", progid);
                results.push((progid.to_string(), true, Some(com_info)));
            }
            Err(e) => {
                log_println!("  ❌ {}: {}", progid, e);
                results.push((progid.to_string(), false, None));
            }
        }
    }
    
    results
}

/// Teste l'instanciation d'un ProgID spécifique
fn test_single_progid(progid: &str) -> SageResult<ComObjectInfo> {
    unsafe {
        // Convertir le ProgID en CLSID
        let progid_hstring = HSTRING::from(progid);
        let clsid = CLSIDFromProgID(&progid_hstring)
            .map_err(|e| errors::SageError::ConversionError {
                from_type: "ProgID".to_string(),
                to_type: "CLSID".to_string(),
                value: format!("Impossible de convertir {}: {}", progid, e),
            })?;
        
        // Essayer de créer l'instance COM
        let dispatch: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_INPROC_SERVER)
            .map_err(|e| errors::SageError::ConversionError {
                from_type: "CLSID".to_string(),
                to_type: "IDispatch".to_string(),
                value: format!("Impossible de créer l'instance {}: {}", progid, e),
            })?;
        
        // Analyser l'objet créé (sans récursion pour éviter les problèmes)
        let com_instance = ComInstance::from_dispatch(dispatch.clone());
        let safe_dispatch = SafeDispatch::new(&dispatch);
        
        let (properties, methods) = analyze_members(&com_instance, &safe_dispatch, 0)?;
        
        Ok(ComObjectInfo {
            name: progid.to_string(),
            object_type: "COM Object (Direct ProgID)".to_string(),
            properties,
            methods,
            sub_objects: Vec::new(), // Pas de récursion pour les objets directs
        })
    }
}

/// Affiche un résumé des Type Libraries découvertes
fn print_typelib_summary(typelibs: &[TypeLibInfo]) {
    log_println!("\n📚 === RÉSUMÉ DES TYPE LIBRARIES ===");
    if typelibs.is_empty() {
        log_println!("❌ Aucune Type Library découverte");
        log_println!("💡 Note: L'analyse des TypeLib nécessite une implémentation plus avancée");
    } else {
        for typelib in typelibs {
            log_println!("📚 {} (v{}) - {} interfaces", 
                typelib.name, typelib.version, typelib.interfaces.len());
        }
    }
}

/// Analyse les objets retournés par certaines méthodes importantes de cpta
fn analyze_cpta_returned_objects(app_dispatch: &IDispatch) -> SageResult<Vec<(String, ComObjectInfo)>> {
    let mut returned_objects = Vec::new();
    
    log_println!("🔍 Analyse des objets retournés par les méthodes...");
    
    let safe_dispatch = SafeDispatch::new(app_dispatch);
    
    // Liste des méthodes importantes à tester
    let important_methods = [
        ("FactoryJournal", "ReadNumero", "VTE"),  // Journal comme premier test
        ("FactoryTiers", "ReadNumero", "BAGUES"), // Client 
        ("FactoryCompteG", "ReadNumero", "601020"), // Compte général
        ("FactoryEcriture", "ReadNumero", "1"), // Écriture
    ];
    
    for (factory_name, method_name, test_param) in &important_methods {
        log_println!("  🧪 Test: {}.{}({})", factory_name, method_name, test_param);
        
        match test_method_return_object(&safe_dispatch, factory_name, method_name, test_param) {
            Ok(Some((object_name, com_info))) => {
                log_println!("  ✅ Objet découvert: {} (type réel)", object_name);
                returned_objects.push((object_name, com_info));
            }
            Ok(None) => {
                log_println!("  ⚠️  Pas d'objet retourné ou objet simple");
            }
            Err(e) => {
                log_println!("  ❌ Erreur: {}", e);
            }
        }
    }
    
    Ok(returned_objects)
}

/// Analyse les objets retournés par certaines méthodes importantes de cial
fn analyze_cial_returned_objects(app_dispatch: &IDispatch) -> SageResult<Vec<(String, ComObjectInfo)>> {
    let mut returned_objects = Vec::new();
    
    log_println!("🔍 Analyse des objets retournés par les méthodes...");
    
    let safe_dispatch = SafeDispatch::new(app_dispatch);
    
    // Liste des méthodes importantes à tester
    let important_methods = [
        ("FactoryArticle", "ReadNumero", "BAAR01"),  // Article comme premier test
        ("FactoryDocument", "ReadNumero", "BC00010"), // Document
        ];
    
    for (factory_name, method_name, test_param) in &important_methods {
        log_println!("  🧪 Test: {}.{}({})", factory_name, method_name, test_param);
        
        match test_method_return_object(&safe_dispatch, factory_name, method_name, test_param) {
            Ok(Some((object_name, com_info))) => {
                log_println!("  ✅ Objet découvert: {} (type réel)", object_name);
                returned_objects.push((object_name, com_info));
            }
            Ok(None) => {
                log_println!("  ⚠️  Pas d'objet retourné ou objet simple");
            }
            Err(e) => {
                log_println!("  ❌ Erreur: {}", e);
            }
        }
    }
    
    Ok(returned_objects)
}

/// Teste le retour d'une méthode pour analyser l'objet COM retourné
fn test_method_return_object(
    app_dispatch: &SafeDispatch,
    factory_name: &str,
    method_name: &str,
    test_param: &str,
) -> SageResult<Option<(String, ComObjectInfo)>> {
    // D'abord récupérer le factory
    let factory_variant = app_dispatch.call_method_by_name(factory_name, &[])?;
    let factory_dispatch = factory_variant.to_dispatch()
        .map_err(|e| errors::SageError::ConversionError {
            from_type: "Variant".to_string(),
            to_type: "IDispatch".to_string(),
            value: format!("Impossible de convertir {} en IDispatch: {}", factory_name, e),
        })?;
    
    let factory_safe = SafeDispatch::new(&factory_dispatch);
    
    // Appeler la méthode avec le paramètre de test
    let param_variant = objets_metier_rs::com::SafeVariant::from_string(test_param);
    let result_variant = factory_safe.call_method_by_name(method_name, &[param_variant])?;
    
    // Vérifier si le résultat est un objet COM
    if result_variant.is_object() {
        if let Ok(result_dispatch) = result_variant.to_dispatch() {
            // Analyser l'objet retourné
            let com_instance = ComInstance::from_dispatch(result_dispatch.clone());
            let result_safe = SafeDispatch::new(&result_dispatch);
            
            let (properties, methods) = analyze_members(&com_instance, &result_safe, 0)?;
            
            let object_name = format!("ReturnedBy.{}.{}", factory_name, method_name);
            let com_info = ComObjectInfo {
                name: object_name.clone(),
                object_type: "COM Object (Returned by method)".to_string(),
                properties,
                methods,
                sub_objects: Vec::new(), // Pas de récursion pour éviter la complexité
            };
            
            return Ok(Some((object_name, com_info)));
        }
    }
    
    Ok(None)
}

/// Affiche un résumé des objets retournés découverts
fn print_returned_objects_summary(returned_objects: &[(String, ComObjectInfo)]) {
    log_println!("\n🔬 === RÉSUMÉ DES OBJETS RETOURNÉS ===");
    if returned_objects.is_empty() {
        log_println!("❌ Aucun objet COM découvert via les retours de méthodes");
        log_println!("💡 Note: Les objets peuvent nécessiter des paramètres spécifiques ou une base ouverte");
    } else {
        log_println!("✅ Objets COM découverts via les retours de méthodes ({}):", returned_objects.len());
        for (name, info) in returned_objects {
            log_println!("  📦 {} - {} propriétés, {} méthodes", 
                name, info.properties.len(), info.methods.len());
        }
    }
}
/// Affiche les résultats des tests de ProgID
fn print_progid_test_results(results: &[(String, bool, Option<ComObjectInfo>)]) {
    log_println!("\n🧪 === RÉSULTATS DES TESTS PROGID ===");
    
    let successful: Vec<_> = results.iter().filter(|(_, success, _)| *success).collect();
    let failed: Vec<_> = results.iter().filter(|(_, success, _)| !*success).collect();
    
    log_println!("✅ Objets COM directement instanciables ({}):", successful.len());
    for (progid, _, info) in &successful {
        if let Some(obj_info) = info {
            log_println!("  📦 {} - {} propriétés, {} méthodes", 
                progid, obj_info.properties.len(), obj_info.methods.len());
        }
    }
    
    log_println!("\n❌ ProgIDs non instanciables ({}):", failed.len());
    for (progid, _, _) in &failed {
        log_println!("  ⚠️  {}", progid);
    }
}

/// Analyse récursivement un objet COM et tous ses sous-objets
fn analyze_com_object_recursive(
    dispatch: &IDispatch,
    object_name: String,
    depth: usize,
    visited_objects: &mut HashSet<String>,
) -> SageResult<ComObjectInfo> {
    let indent = "  ".repeat(depth);
    log_println!("{}🔍 Analyse de: {}", indent, object_name);
    
    // Éviter les boucles infinies
    if visited_objects.contains(&object_name) {
        log_println!("{}⚠️  Objet déjà visité, évitement de boucle", indent);
        return Ok(ComObjectInfo {
            name: object_name,
            object_type: "Previously analyzed".to_string(),
            properties: vec![],
            methods: vec![],
            sub_objects: vec![],
        });
    }
    visited_objects.insert(object_name.clone());
    
    let safe_dispatch = SafeDispatch::new(dispatch);
    let com_instance = ComInstance::from_dispatch(dispatch.clone());
    
    // Analyser les méthodes et propriétés
    let (properties, methods) = analyze_members(&com_instance, &safe_dispatch, depth)?;
    
    // Identifier les sous-objets et les analyser récursivement
    let mut sub_objects = Vec::new();
    
    for property in &properties {
        if property.is_object && depth < 5 { // Limiter la profondeur pour éviter l'explosion
            log_println!("{}🎯 Analyse récursive de la propriété objet: {}", indent, property.name);
            
            match safe_dispatch.call_method_by_name(&property.name, &[]) {
                Ok(sub_object_variant) => {
                    if let Ok(sub_dispatch) = sub_object_variant.to_dispatch() {
                        let sub_object_name = format!("{}.{}", object_name, property.name);
                        
                        match analyze_com_object_recursive(
                            &sub_dispatch, 
                            sub_object_name,
                            depth + 1,
                            visited_objects
                        ) {
                            Ok(sub_info) => sub_objects.push(sub_info),
                            Err(e) => {
                                log_println!("{}❌ Erreur analyse sous-objet {}: {}", indent, property.name, e);
                            }
                        }
                    }
                }
                Err(e) => {
                    log_println!("{}⚠️  Impossible d'accéder à {}: {}", indent, property.name, e);
                }
            }
        }
    }
    
    log_println!("{}✅ Analyse terminée: {} propriétés, {} méthodes, {} sous-objets", 
        indent, properties.len(), methods.len(), sub_objects.len());
    
    Ok(ComObjectInfo {
        name: object_name,
        object_type: "COM Object".to_string(),
        properties,
        methods,
        sub_objects,
    })
}

/// Analyse les membres (propriétés et méthodes) d'un objet COM
fn analyze_members(
    instance: &ComInstance,
    safe_dispatch: &SafeDispatch,
    depth: usize,
) -> SageResult<(Vec<PropertyInfo>, Vec<MethodInfo>)> {
    let indent = "  ".repeat(depth);
    
    // Analyser les propriétés groupées
    let mut properties = Vec::new();
    if let Ok(grouped_props) = instance.group_properties() {
        for (name, variants) in grouped_props {
            let access_types: Vec<String> = variants.iter().map(|v| {
                match v.member_type {
                    MemberType::PropertyGet => "get".to_string(),
                    MemberType::PropertyPut => "put".to_string(),
                    MemberType::PropertyPutRef => "putref".to_string(),
                    _ => "?".to_string(),
                }
            }).collect();
            
            let return_type = variants.first()
                .and_then(|v| v.return_type.as_deref())
                .unwrap_or("Unknown");
            
            let id = variants.first().map_or(0, |v| v.id);
            
            // Tester si c'est un objet en essayant d'y accéder
            let is_object = test_if_property_is_object(safe_dispatch, &name);
            
            // Créer une description enrichie
            let description = if is_object {
                Some(format!("Objet COM {} - Accès: {}", return_type, access_types.join("/")))
            } else {
                Some(format!("Propriété {} - Accès: {}", return_type, access_types.join("/")))
            };
            
            properties.push(PropertyInfo {
                name: name.clone(),
                property_type: return_type.to_string(),
                is_object,
                access_type: access_types.join("/"),
                id,
                description,
            });
            
            if is_object {
                log_println!("{}📦 Propriété objet détectée: {}", indent, name);
            }
        }
    }
    
    // Analyser les méthodes
    let mut methods = Vec::new();
    if let Ok(method_list) = instance.list_methods_only() {
        for method in method_list {
            // Essayer d'extraire plus d'informations sur les paramètres
            let parameters = extract_method_parameters(instance, &method.name, method.param_count);
            
            methods.push(MethodInfo {
                name: method.name.clone(),
                return_type: method.return_type.unwrap_or_else(|| "Unknown".to_string()),
                parameter_count: method.param_count,
                id: method.id,
                parameters,
            });
        }
    }
    
    Ok((properties, methods))
}

/// Teste si une propriété retourne un objet COM
fn test_if_property_is_object(safe_dispatch: &SafeDispatch, property_name: &str) -> bool {
    // Liste des propriétés connues qui sont des objets
    let known_objects = [
        "Loggable", "FactoryJournal", "FactoryTiers", "FactoryArticle", 
        "FactoryCompteG", "FactoryEcritureC", "FactoryLivraison", "FactoryDocument",
        "FactoryDevis", "FactoryFacture", "FactoryStock", "FactoryInventaire",
        "DocumentsVente", "DocumentsAchat", "DocumentsStock", "Parametre"
    ];
    
    if known_objects.contains(&property_name) {
        return true;
    }
    
    // Tester dynamiquement
    match safe_dispatch.call_method_by_name(property_name, &[]) {
        Ok(variant) => variant.is_object(),
        Err(_) => false,
    }
}

/// Extrait les informations sur les paramètres d'une méthode
fn extract_method_parameters(
    _instance: &ComInstance,
    _method_name: &str,
    param_count: Option<u32>,
) -> Vec<ParameterInfo> {
    // Pour l'instant, nous créons des paramètres génériques
    // Une version plus avancée pourrait utiliser ITypeInfo pour obtenir les vrais types
    match param_count {
        Some(count) => {
            (0..count).map(|i| ParameterInfo {
                name: Some(format!("param{}", i)),
                param_type: "VARIANT".to_string(), // Type générique pour l'instant
                is_optional: false,
            }).collect()
        }
        None => vec![],
    }
}

/// Affiche un résumé de l'analyse
fn print_analysis_summary(root: &ComObjectInfo) {
    log_println!("\n📊 === RÉSUMÉ DE L'ANALYSE ===");
    print_object_summary(root, 0);
}

/// Affiche récursivement le résumé d'un objet
fn print_object_summary(obj: &ComObjectInfo, depth: usize) {
    let indent = "  ".repeat(depth);
    log_println!("{}📦 {}", indent, obj.name);
    log_println!("{}   📋 {} propriétés, 🔧 {} méthodes", 
        indent, obj.properties.len(), obj.methods.len());
    
    // Afficher les propriétés avec détails
    if !obj.properties.is_empty() && depth < 3 { // Limiter les détails pour éviter le spam
        log_println!("{}   📋 PROPRIÉTÉS DÉTAILLÉES:", indent);
        for prop in &obj.properties {
            let object_marker = if prop.is_object { " [OBJET]" } else { "" };
            let type_info = format!("{} [{}]", prop.property_type, prop.access_type);
            log_println!("{}      [{}] {} : {}{}", 
                indent, prop.id, prop.name, type_info, object_marker);
            if let Some(desc) = &prop.description {
                log_println!("{}           {}", indent, desc);
            }
        }
    }
    
    // Afficher les méthodes avec détails
    if !obj.methods.is_empty() && depth < 3 { // Limiter les détails pour éviter le spam
        log_println!("{}   🔧 MÉTHODES DÉTAILLÉES:", indent);
        for method in &obj.methods {
            let param_info = if method.parameters.is_empty() {
                "()".to_string()
            } else {
                let params: Vec<String> = method.parameters.iter().map(|p| {
                    format!("{}: {}", 
                        p.name.as_deref().unwrap_or("param"), 
                        p.param_type)
                }).collect();
                format!("({})", params.join(", "))
            };
            log_println!("{}      [{}] {}{} -> {}", 
                indent, method.id, method.name, param_info, method.return_type);
        }
    }
    
    // Afficher les sous-objets récursivement
    for sub_obj in &obj.sub_objects {
        print_object_summary(sub_obj, depth + 1);
    }
}
