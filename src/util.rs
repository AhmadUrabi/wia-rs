use std::mem::ManuallyDrop;

use windows::{
    Win32::{
        Devices::ImageAcquisition::*,
        System::{
            Com::{
                StructuredStorage::{
                    PROPSPEC, PROPSPEC_0, PROPVARIANT, PRSPEC_PROPID, PropVariantClear,
                },
            },
            Variant::*,
        },
    },
    core::*,
};

pub(crate) fn read_bstr_property(prop_storage: &IWiaPropertyStorage, prop_id: u32) -> std::result::Result<String, String> {
    unsafe {
        let mut property_id = PROPSPEC {
            ulKind: PRSPEC_PROPID,
            Anonymous: PROPSPEC_0 { propid: prop_id },
        };
        let mut property_variant = PROPVARIANT::default();

        prop_storage.ReadMultiple(1, &mut property_id, &mut property_variant).map_err(|e| handle_error(e))?;

        let result = if property_variant.vt() == VT_BSTR {
            let bstr = ManuallyDrop::into_inner(
                property_variant
                    .Anonymous
                    .Anonymous
                    .Anonymous
                    .bstrVal
                    .clone(),
            );
            if !bstr.is_empty() {
                let s = bstr.to_string();
                // Don't free the BSTR - PropVariantClear will do that
                std::mem::forget(bstr);
                s
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        PropVariantClear(&mut property_variant).map_err(|e| handle_error(e))?;
        Ok(result)
    }
}

const ERROR_CODES: [(&str, (&str, &str)); 20] = [
    (
        "0x80210006",
        (
            "WIA_ERROR_BUSY",
            "The device is busy. Close any apps that are using this device or wait for it to finish and then try again.",
        ),
    ),
    (
        "0x80210016",
        (
            "WIA_ERROR_COVER_OPEN",
            "One or more of the device’s cover is open.",
        ),
    ),
    (
        "0x8021000A",
        (
            "WIA_ERROR_DEVICE_COMMUNICATION",
            "Communication with the WIA device failed. Make sure that the device is powered on and connected to the PC. If the problem persists, disconnect and reconnect the device.",
        ),
    ),
    (
        "0x8021000D",
        (
            "WIA_ERROR_DEVICE_LOCKED",
            "The device is locked. Close any apps that are using this device or wait for it to finish and then try again.",
        ),
    ),
    (
        "0x8021000E",
        (
            "WIA_ERROR_EXCEPTION_IN_DRIVER",
            "The device driver threw an exception.",
        ),
    ),
    (
        "0x80210001",
        (
            "WIA_ERROR_GENERAL_ERROR",
            "An unknown error has occurred with the WIA device.",
        ),
    ),
    (
        "0x8021000C",
        (
            "WIA_ERROR_INCORRECT_HARDWARE_SETTING",
            "There is an incorrect setting on the WIA device.",
        ),
    ),
    (
        "0x8021000B",
        (
            "WIA_ERROR_INVALID_COMMAND",
            "The device doesn't support this command.",
        ),
    ),
    (
        "0x8021000F",
        (
            "WIA_ERROR_INVALID_DRIVER_RESPONSE",
            "The response from the driver is invalid.",
        ),
    ),
    (
        "0x80210009",
        (
            "WIA_ERROR_ITEM_DELETED",
            "The WIA device was deleted. It's no longer available.",
        ),
    ),
    (
        "0x80210017",
        ("WIA_ERROR_LAMP_OFF", "The scanner's lamp is off."),
    ),
    (
        "0x80210021",
        (
            "WIA_ERROR_MAXIMUM_PRINTER_ENDORSER_COUNTER",
            "A scan job was interrupted because an Imprinter/Endorser item reached the maximum valid value for WIA_IPS_PRINTER_ENDORSER_COUNTER, and was reset to 0. This feature is available with Windows 8 and later versions of Windows.",
        ),
    ),
    (
        "0x80210020",
        (
            "WIA_ERROR_MULTI_FEED",
            "A scan error occurred because of a multiple page feed condition. This feature is available with Windows 8 and later versions of Windows.",
        ),
    ),
    (
        "0x80210005",
        (
            "WIA_ERROR_OFFLINE",
            "The device is offline. Make sure the device is powered on and connected to the PC.",
        ),
    ),
    (
        "0x80210003",
        (
            "WIA_ERROR_PAPER_EMPTY",
            "There are no documents in the document feeder.",
        ),
    ),
    (
        "0x80210002",
        (
            "WIA_ERROR_PAPER_JAM",
            "Paper is jammed in the scanner's document feeder.",
        ),
    ),
    (
        "0x80210004",
        (
            "WIA_ERROR_PAPER_PROBLEM",
            "An unspecified problem occurred with the scanner's document feeder.",
        ),
    ),
    (
        "0x80210007",
        ("WIA_ERROR_WARMING_UP", "The device is warming up."),
    ),
    (
        "0x80210008",
        (
            "WIA_ERROR_USER_INTERVENTION",
            "There is a problem with the WIA device. Make sure that the device is turned on, online, and any cables are properly connected.",
        ),
    ),
    (
        "0x80210015",
        (
            "WIA_S_NO_DEVICE_AVAILABLE",
            "No scanner device was found. Make sure the device is online, connected to the PC, and has the correct driver installed on the PC.",
        ),
    ),
];


pub(crate) fn get_error(error_code: &str) -> Option<(&str, &str)> {
    ERROR_CODES
        .iter()
        .find(|(code, _)| *code == error_code)
        .map(|(_, (name, desc))| (*name, *desc))
}

pub(crate) fn handle_error(err: Error) -> String {
    let binding = err.code().to_string();
    let code = binding.as_str();
    let data = get_error(code);
    if let Some((name, desc)) = data {
        return format!("{} - {} - {}", code, name, desc);
    } else {
        return format!("Unknown error");
    }
}
