# Recherche VARIANT - Implémentation COM/Rust avec windows-rs

## Vue d'ensemble

Ce document compile les informations essentielles pour l'implémentation complète de la conversion VARIANT dans notre projet v0.1.3. Les VARIANT sont la base de l'interopérabilité COM en Windows.

## 1. Structure VARIANT dans windows-rs

### Structure de base
```rust
#[repr(C)]
pub struct VARIANT {
    pub Anonymous: VARIANT_0,
}

#[repr(C)]
pub union VARIANT_0 {
    pub Anonymous: core::mem::ManuallyDrop<VARIANT_0_0>,
    pub decVal: super::super::Foundation::DECIMAL,
}

#[repr(C)]
pub struct VARIANT_0_0 {
    pub vt: VARENUM,           // Type de données
    pub wReserved1: u16,
    pub wReserved2: u16,
    pub wReserved3: u16,
    pub Anonymous: VARIANT_0_0_0,
}

#[repr(C)]
pub union VARIANT_0_0_0 {
    pub llVal: i64,
    pub lVal: i32,
    pub bVal: u8,
    pub iVal: i16,
    pub fltVal: f32,
    pub dblVal: f64,
    pub boolVal: super::super::Foundation::VARIANT_BOOL,
    pub scode: i32,
    pub cyVal: super::Com::CY,
    pub date: f64,
    pub bstrVal: core::mem::ManuallyDrop<windows_core::BSTR>,
    pub punkVal: core::mem::ManuallyDrop<Option<windows_core::IUnknown>>,
    pub pdispVal: core::mem::ManuallyDrop<Option<super::Com::IDispatch>>,
    pub parray: *mut super::Com::SAFEARRAY,
    // ... autres champs
}
```

## 2. Types VARENUM essentiels

```rust
pub const VT_EMPTY: VARENUM = VARENUM(0u16);
pub const VT_NULL: VARENUM = VARENUM(1u16);
pub const VT_I2: VARENUM = VARENUM(2u16);      // i16
pub const VT_I4: VARENUM = VARENUM(3u16);      // i32
pub const VT_R4: VARENUM = VARENUM(4u16);      // f32
pub const VT_R8: VARENUM = VARENUM(5u16);      // f64
pub const VT_CY: VARENUM = VARENUM(6u16);      // Currency
pub const VT_DATE: VARENUM = VARENUM(7u16);    // Date (f64)
pub const VT_BSTR: VARENUM = VARENUM(8u16);    // String
pub const VT_BOOL: VARENUM = VARENUM(11u16);   // Boolean
pub const VT_VARIANT: VARENUM = VARENUM(12u16);
pub const VT_I1: VARENUM = VARENUM(16u16);     // i8
pub const VT_UI1: VARENUM = VARENUM(17u16);    // u8
pub const VT_UI2: VARENUM = VARENUM(18u16);    // u16
pub const VT_UI4: VARENUM = VARENUM(19u16);    // u32
pub const VT_I8: VARENUM = VARENUM(20u16);     // i64
pub const VT_UI8: VARENUM = VARENUM(21u16);    // u64
```

## 3. VARIANT_BOOL - Point critique

### Implémentation correcte
```rust
// Dans windows-rs/src/extensions/Win32/Foundation/VARIANT_BOOL.rs
impl VARIANT_BOOL {
    #[inline]
    pub fn as_bool(self) -> bool {
        self.0 != 0
    }
}

impl From<bool> for VARIANT_BOOL {
    fn from(value: bool) -> Self {
        if value {
            VARIANT_TRUE    // -1i16
        } else {
            VARIANT_FALSE   // 0i16
        }
    }
}
```

### Constantes importantes
```rust
pub const VARIANT_TRUE: VARIANT_BOOL = VARIANT_BOOL(-1i16);
pub const VARIANT_FALSE: VARIANT_BOOL = VARIANT_BOOL(0i16);
```

## 4. Fonctions de conversion Windows API

### Fonctions disponibles dans windows-rs
```rust
// Conversion de type automatique
pub unsafe fn VariantChangeType(
    pvargdest: *mut VARIANT, 
    pvarsrc: *const VARIANT, 
    wflags: VAR_CHANGE_FLAGS, 
    vt: VARENUM
) -> windows_core::Result<()>

// Nettoyage mémoire
pub unsafe fn VariantClear(pvarg: *mut VARIANT) -> windows_core::Result<()>

// Copie
pub unsafe fn VariantCopy(
    pvargdest: *mut VARIANT, 
    pvargsrc: *const VARIANT
) -> windows_core::Result<()>

// Conversion vers types spécifiques
pub unsafe fn VariantToBoolean(varin: *const VARIANT) -> windows_core::Result<windows_core::BOOL>
pub unsafe fn VariantToString(varin: *const VARIANT, pwszBuf: &mut [u16]) -> windows_core::Result<()>
```

### Flags de conversion
```rust
pub const VARIANT_ALPHABOOL: VAR_CHANGE_FLAGS = VAR_CHANGE_FLAGS(2u16);
pub const VARIANT_LOCALBOOL: VAR_CHANGE_FLAGS = VAR_CHANGE_FLAGS(16u16);
pub const VARIANT_NOUSEROVERRIDE: VAR_CHANGE_FLAGS = VAR_CHANGE_FLAGS(4u16);
```

## 5. Gestion mémoire et sécurité

### ManuallyDrop pour les types complexes
```rust
// BSTR et objets COM utilisent ManuallyDrop
pub bstrVal: core::mem::ManuallyDrop<windows_core::BSTR>,
pub punkVal: core::mem::ManuallyDrop<Option<windows_core::IUnknown>>,
```

### Utilisation de std::ptr::write pour les unions
```rust
// Éviter les erreurs de déréférencement automatique
unsafe {
    std::ptr::write(&mut variant.Anonymous.Anonymous.vt, VT_BSTR);
    std::ptr::write(&mut variant.Anonymous.Anonymous.Anonymous.bstrVal, 
                   ManuallyDrop::new(bstr));
}
```

## 6. Exemples d'implémentation windows-rs

### Extensions VARIANT dans windows-rs
```rust
// Conversion automatique depuis différents types
impl From<&str> for VARIANT {
    fn from(value: &str) -> Self {
        VARIANT::from(BSTR::from(value))
    }
}

impl From<bool> for VARIANT {
    fn from(value: bool) -> Self {
        let mut variant = VARIANT::default();
        variant.Anonymous.Anonymous.vt = VT_BOOL;
        variant.Anonymous.Anonymous.Anonymous.boolVal = value.into();
        variant
    }
}
```

## 7. Patterns d'utilisation recommandés

### Création sécurisée de VARIANT
```rust
pub fn create_bstr_variant(s: &str) -> SageResult<VARIANT> {
    unsafe {
        let mut variant = VARIANT::default();
        let bstr = BSTR::from(s);
        
        std::ptr::write(&mut variant.Anonymous.Anonymous.vt, VT_BSTR);
        std::ptr::write(&mut variant.Anonymous.Anonymous.Anonymous.bstrVal, 
                       ManuallyDrop::new(bstr));
        
        Ok(variant)
    }
}
```

### Lecture sécurisée de VARIANT
```rust
pub fn read_variant_safely(variant: &VARIANT) -> SageResult<SafeVariant> {
    unsafe {
        match variant.Anonymous.Anonymous.vt {
            VT_BSTR => {
                let bstr = &variant.Anonymous.Anonymous.Anonymous.bstrVal;
                if bstr.is_empty() {
                    Ok(SafeVariant::BStr(String::new()))
                } else {
                    let rust_string = bstr.to_string();
                    Ok(SafeVariant::BStr(rust_string))
                }
            },
            VT_BOOL => {
                let bool_val = variant.Anonymous.Anonymous.Anonymous.boolVal;
                Ok(SafeVariant::Bool(bool_val.as_bool()))
            },
            // ... autres types
            _ => Err(SageError::ConversionError {
                from_type: format!("VT_{}", variant.Anonymous.Anonymous.vt.0),
                to_type: "SafeVariant".to_string(),
                value: "Type non supporté".to_string(),
            })
        }
    }
}
```

## 8. Problèmes identifiés et solutions

### ✅ VARIANT_BOOL - Solution trouvée
```rust
// Problème: VARIANT_BOOL non trouvé
// Solution: Import correct et conversion manuelle
use windows::Win32::Foundation::{VARIANT_BOOL, VARIANT_TRUE, VARIANT_FALSE};

let bool_val: VARIANT_BOOL = if my_bool { VARIANT_TRUE } else { VARIANT_FALSE };
```

### ✅ Accès aux champs union - Solution
```rust
// Problème: Erreur ManuallyDrop sur union
// Solution: Utiliser std::ptr::write
unsafe {
    std::ptr::write(&mut variant.Anonymous.Anonymous.vt, VT_BSTR);
    std::ptr::write(&mut variant.Anonymous.Anonymous.Anonymous.bstrVal, 
                   ManuallyDrop::new(bstr));
}
```

### ✅ VariantChangeType - Usage correct
```rust
// Problème: VAR_CHANGE_FLAGS attendu au lieu de u16
// Solution: Wrapper correct
unsafe {
    VariantChangeType(
        &mut dest_variant,
        &src_variant,
        VAR_CHANGE_FLAGS(0),  // Pas juste 0
        VT_BSTR
    )?;
}
```

## 9. Plan d'implémentation v0.1.3

### Phase 1: Types de base (PRIORITÉ HAUTE)
- [x] VT_EMPTY, VT_NULL
- [x] VT_BSTR (chaînes) 
- [x] VT_BOOL (booléens)
- [x] VT_I4 (entiers 32-bit)
- [x] VT_R8 (flottants 64-bit)

### Phase 2: Types numériques étendus
- [ ] VT_I2, VT_I8, VT_UI1, VT_UI2, VT_UI4, VT_UI8
- [ ] VT_R4 (float 32-bit)
- [ ] VT_CY (currency)
- [ ] VT_DATE (dates)

### Phase 3: Types avancés (v0.1.4)
- [ ] VT_DISPATCH (objets COM)
- [ ] VT_ARRAY (tableaux)
- [ ] VT_VARIANT (variants imbriqués)

## 10. Imports nécessaires

```rust
use windows::{
    Win32::System::Variant::*,
    Win32::Foundation::{VARIANT_BOOL, VARIANT_TRUE, VARIANT_FALSE},
    core::*,
};
use std::mem::ManuallyDrop;
```

## 11. Tests de validation

```rust
#[cfg(test)]
mod variant_tests {
    use super::*;

    #[test]
    fn test_bstr_conversion() {
        let variant = SafeVariant::from_string("Test");
        let win_variant = variant.to_variant().unwrap();
        let back = SafeVariant::from_variant(win_variant).unwrap();
        
        match back {
            SafeVariant::BStr(s) => assert_eq!(s, "Test"),
            _ => panic!("Conversion échouée"),
        }
    }

    #[test]
    fn test_bool_conversion() {
        let variant = SafeVariant::from_bool(true);
        let win_variant = variant.to_variant().unwrap();
        let back = SafeVariant::from_variant(win_variant).unwrap();
        
        match back {
            SafeVariant::Bool(b) => assert_eq!(b, true),
            _ => panic!("Conversion booléenne échouée"),
        }
    }
}
```

---

**Note importante**: Cette recherche forme la base technique pour résoudre l'erreur "Conversion VARIANT non implémentée (BStr)" et permettre les appels COM fonctionnels avec paramètres dans v0.1.3.
