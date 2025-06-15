use std::collections::HashMap;
use windows::{
    Win32::{
        Devices::ImageAcquisition::*,
        System::{
            Com::{
                StructuredStorage::{
                    PROPSPEC, PROPSPEC_0, PROPVARIANT, PRSPEC_PROPID, PropVariantClear,
                },
                *,
            },
            Variant::*,
        },
    },
    core::*,
};

mod util;

use util::{read_bstr_property, handle_error};

struct WIAScanManager;

impl Drop for WIAScanManager {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

impl WIAScanManager {
    pub fn init() -> std::result::Result<Self, String> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED).unwrap();
        }

        Ok(WIAScanManager)
    }

    pub fn list_devices() -> std::result::Result<(), String> {
        println!("Scanning for WIA devices...");

        // Create a collection to store device IDs
        let mut device_map: HashMap<usize, String> = HashMap::new();

        // Create WIA device manager
        unsafe {
            let device_manager: IWiaDevMgr =
                CoCreateInstance(&WiaDevMgr, None, CLSCTX_LOCAL_SERVER).unwrap();

            // Enumerate WIA devices
            let enum_wia_dev: Option<IEnumWIA_DEV_INFO> = device_manager
                .EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL as i32)
                .ok();

            if enum_wia_dev.is_none() {
                println!("No WIA devices found.");
                return Ok(());
            }
            let enum_wia_dev = enum_wia_dev.unwrap();

            // Get device count
            let device_count = enum_wia_dev.GetCount().map_err(|e| handle_error(e))?;
            println!("Found {} WIA device(s)", device_count);

            // Iterate through devices
            for i in 0..device_count {
                // Get device info
                let mut wia_dev_info: Option<IWiaPropertyStorage> = None;
                enum_wia_dev
                    .Next(1, &mut wia_dev_info, std::ptr::null_mut())
                    .map_err(|e| handle_error(e))?;

                if let Some(dev_info) = wia_dev_info {
                    // Get device ID
                    let id_string = read_bstr_property(&dev_info, WIA_DIP_DEV_ID)?;
                    println!("Device {}: ID = {}", i + 1, id_string);

                    // Store device ID in map
                    device_map.insert((i + 1).try_into().unwrap(), id_string);

                    // Get device name
                    let name_string = read_bstr_property(&dev_info, WIA_DIP_DEV_NAME)?;
                    println!("      Name = {}", name_string);

                    // Get device description
                    let desc_string = read_bstr_property(&dev_info, WIA_DIP_DEV_DESC)?;
                    println!("      Description = {}", desc_string);

                    println!(); // Add empty line between devices
                }
            }

            // Check if any devices were found
            if !device_map.is_empty() {
                println!("Would you like to scan a document? (y/n)");
                let mut input = String::new();
                std::io::stdin().read_line(&mut input).unwrap();

                if input.trim().to_lowercase() == "y" {
                    println!("Enter the device number to use for scanning:");
                    input.clear();
                    std::io::stdin().read_line(&mut input).unwrap();

                    if let Ok(device_num) = input.trim().parse::<usize>() {
                        if let Some(device_id) = device_map.get(&device_num) {
                            // Create device manager and device to check capabilities
                            let device_manager: IWiaDevMgr =
                                CoCreateInstance(&WiaDevMgr, None, CLSCTX_LOCAL_SERVER)
                                    .map_err(|e| handle_error(e))?;
                            println!("Device Manager created successfully");
                            let device: IWiaItem = device_manager
                                .CreateDevice(&BSTR::from(device_id))
                                .map_err(|e| handle_error(e))?;
                            println!("Device created successfully");

                            // Find the scanner item
                            let enum_items: IEnumWiaItem =
                                device.EnumChildItems().map_err(|e| handle_error(e))?;
                            let mut scanner_item: Option<IWiaItem> = None;
                            let mut num_fetched: u32 = 0;
                            enum_items
                                .Next(1, &mut scanner_item, &mut num_fetched)
                                .map_err(|e| handle_error(e))?;
                            println!("Scanner item found successfully");

                            if let Some(item) = scanner_item {
                                let props: IWiaPropertyStorage =
                                    item.cast().map_err(|e| handle_error(e))?;
                                // First check device level properties for capability detection
                                println!("Checking device level properties...");
                                let device_props: IWiaPropertyStorage =
                                    device.cast().map_err(|e| handle_error(e))?;
                                let (has_feeder_device, has_flatbed_device) =
                                    Self::check_scanner_capabilities(&device_props)?;

                                // Then check item level properties
                                println!("Checking item level properties...");
                                let (has_feeder_item, has_flatbed_item) =
                                    Self::check_scanner_capabilities(&props)?;

                                // Combine results - if either level reports capability, consider it available
                                let has_feeder = has_feeder_device || has_feeder_item;
                                let has_flatbed = has_flatbed_device || has_flatbed_item;

                                println!(
                                    "Final capability detection: Feeder: {}, Flatbed: {}",
                                    has_feeder, has_flatbed
                                );

                                let use_feeder;

                                if has_feeder && has_flatbed {
                                    println!("Select scan source:");
                                    println!("1. Flatbed");
                                    println!("2. Document Feeder");
                                    input.clear();
                                    std::io::stdin().read_line(&mut input).unwrap();

                                    use_feeder = match input.trim() {
                                        "2" => true,
                                        _ => false, // Default to flatbed for any other input
                                    };
                                } else if has_feeder {
                                    println!("Only document feeder available. Using feeder.");
                                    use_feeder = true;
                                } else {
                                    println!("Only flatbed available. Using flatbed.");
                                    use_feeder = false;
                                }

                                println!(
                                    "Starting scan with {} source...",
                                    if use_feeder { "feeder" } else { "flatbed" }
                                );
                                Self::scan_document(device_id, use_feeder)?;
                            } else {
                                println!("No scanner item found");
                            }
                        } else {
                            println!("Invalid device number.");
                        }
                    } else {
                        println!("No scanner device found");
                    }
                }
            }
        }


        Ok(())
    }

    // Function to check scanner capabilities
    fn check_scanner_capabilities(
        props: &IWiaPropertyStorage,
    ) -> std::result::Result<(bool, bool), String> {
        unsafe {
            // Check document handling capabilities
            let mut prop_id = PROPSPEC {
                ulKind: PRSPEC_PROPID,
                Anonymous: PROPSPEC_0 {
                    propid: WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES,
                },
            };
            let mut prop_var = PROPVARIANT::default();

            let hr = props.ReadMultiple(1, &mut prop_id, &mut prop_var);

            let mut has_feeder = false;
            let mut has_flatbed = false;

            println!("Checking scanner capabilities...");

            if hr.is_ok() {
                if prop_var.vt() == VT_I4 {
                    let capabilities = prop_var.Anonymous.Anonymous.Anonymous.lVal;
                    println!("Capabilities value: {}", capabilities);

                    // Debug specific capability flags
                    println!("FEEDER value: {}", FEEDER as i32);
                    println!("FLATBED value: {}", FLATBED as i32);

                    has_feeder = (capabilities & (FEEDER as i32)) != 0;
                    has_flatbed = (capabilities & (FLATBED as i32)) != 0;

                    println!("Has feeder: {}, Has flatbed: {}", has_feeder, has_flatbed);
                } else {
                    // println!("Unexpected property type: {} (expected VT_I4)", prop_var.vt.0);
                }

                PropVariantClear(&mut prop_var).map_err(|e| handle_error(e))?;
            } else {
                println!("Failed to read document handling capabilities: {:?}", hr);

                // Try to read device properties directly
                let mut prop_status = PROPSPEC {
                    ulKind: PRSPEC_PROPID,
                    Anonymous: PROPSPEC_0 {
                        propid: WIA_DPS_DOCUMENT_HANDLING_STATUS,
                    },
                };
                let mut status_var = PROPVARIANT::default();
                if props
                    .ReadMultiple(1, &mut prop_status, &mut status_var)
                    .is_ok()
                {
                    if status_var.vt() == VT_I4 {
                        let status = status_var.Anonymous.Anonymous.Anonymous.lVal;
                        println!("Document handling status: {}", status);
                        has_feeder = (status & (FEEDER as i32)) != 0;
                        has_flatbed = true; // Assume flatbed is available
                    }
                    PropVariantClear(&mut status_var).map_err(|e| handle_error(e))?;
                }
            }

            // Default to assuming both are available if detection fails
            if !has_feeder && !has_flatbed {
                println!("Could not detect capabilities, assuming both are available");
                has_feeder = true;
                has_flatbed = true;
            }

            Ok((has_feeder, has_flatbed))
        }
    }

    pub fn scan_document(device_id: &str, use_feeder: bool) -> std::result::Result<(), String> {
        println!("Scanning document from device: {}", device_id);
        unsafe {
            // Create WIA device manager
            let device_manager: IWiaDevMgr =
                CoCreateInstance(&WiaDevMgr, None, CLSCTX_LOCAL_SERVER).map_err(|e| handle_error(e))?;
            println!("WIA Device Manager created.");

            // Connect to the specific device
            println!("Connecting to device: {}", device_id);
            let device: IWiaItem = device_manager
                .CreateDevice(&BSTR::from(device_id))
                .map_err(|e| handle_error(e))?;
            println!("Connected to device: {}", device_id);

            // Set document handling on the root device
            let device_props: IWiaPropertyStorage = device.cast().map_err(|e| handle_error(e))?;
            let mut prop_id = PROPSPEC {
                ulKind: PRSPEC_PROPID,
                Anonymous: PROPSPEC_0 {
                    propid: WIA_IPS_DOCUMENT_HANDLING_SELECT,
                },
            };
            let mut prop_var = PROPVARIANT::default();
            let handling_value = if use_feeder { FEEDER } else { FLATBED };
            println!(
                "Setting document handling select to: {} ({})",
                if use_feeder { "FEEDER" } else { "FLATBED" },
                handling_value as i32
            );
            (*prop_var.Anonymous.Anonymous).Anonymous.lVal = handling_value as i32;
            (*prop_var.Anonymous.Anonymous).vt = VT_I4;
            let hr = device_props.WriteMultiple(1, &mut prop_id, &mut prop_var, 1);
            if hr.is_err() {
                println!("Warning: Failed to set document handling mode: {:?}", hr);
                // Try to continue anyway
            }

            // Re-enumerate to get the correct scanning item
            let enum_items: IEnumWiaItem = device.EnumChildItems().map_err(|e| handle_error(e))?;
            let mut scan_item: Option<IWiaItem> = None;
            let mut num_fetched: u32 = 0;
            enum_items
                .Next(1, &mut scan_item, &mut num_fetched)
                .map_err(|e| handle_error(e))?;
            if scan_item.is_none() {
                println!("No scan item found after setting handling mode.");
                return Ok(());
            }
            let scan_item = scan_item.unwrap();

            // Create a temporary file path for the output
            let output_path = "scanned_document.pdf";
            let wide_path: Vec<u16> = output_path
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect();

            // Set up the transfer medium
            let mut stgm = STGMEDIUM::default();
            stgm.tymed = TYMED_FILE.0 as u32;
            stgm.u.lpszFileName = PWSTR(wide_path.as_ptr() as *mut u16);

            // Get the IWiaDataTransfer from the scan item
            let data_transfer: IWiaDataTransfer = scan_item.cast().map_err(|e| handle_error(e))?;

            println!("Saving document to {}", output_path);
            data_transfer
                .idtGetData(&mut stgm, None)
                .map_err(|e| handle_error(e))?;

            println!("Scan complete! Document saved as: {}", output_path);
            Ok(())
        }
    }

}
