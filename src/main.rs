use std::io::Write;

use windows::{
    core::*,
    Win32::{
        NetworkManagement::WindowsFirewall::{IStaticPortMapping, IUPnPNAT, UPnPNAT},
        System::{
            Com::{
                CoCreateInstance, CoInitialize, CoUninitialize, CLSCTX_INPROC_HANDLER,
                CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER,
            },
            Ole::IEnumVARIANT,
            Variant::{VARIANT, VT_DISPATCH},
        },
    },
};

fn main() -> Result<()> {
    unsafe {
        CoInitialize(None)?;
        let upnp_nat = CoCreateInstance::<_, IUPnPNAT>(
            &UPnPNAT,
            None,
            CLSCTX_INPROC_SERVER
                | CLSCTX_INPROC_HANDLER
                | CLSCTX_LOCAL_SERVER
                | CLSCTX_REMOTE_SERVER,
        )?;
        let static_port_mapping_collection = upnp_nat.StaticPortMappingCollection()?;
        let unknown = static_port_mapping_collection._NewEnum()?;
        let enum_variant = unknown.cast::<IEnumVARIANT>()?;
        enum_variant.Reset()?;
        let mut variant = std::iter::repeat(VARIANT::default())
            .take(static_port_mapping_collection.Count()? as usize)
            .collect::<Vec<VARIANT>>();
        let mut count = 0;
        enum_variant.Next(&mut variant, &mut count).ok()?;
        let enum_static_port_mapping = variant
            .into_iter()
            .take(count as usize)
            .filter_map(|x| {
                (x.Anonymous.Anonymous.vt == VT_DISPATCH).then(|| {
                    x.Anonymous
                        .Anonymous
                        .Anonymous
                        .pdispVal
                        .as_ref()
                        .unwrap()
                        .cast::<IStaticPortMapping>()
                        .unwrap()
                })
            })
            .map(|x| StaticPortMapping {
                external_port: x.clone().ExternalPort().unwrap(),
                internal_port: x.clone().InternalPort().unwrap(),
                protocol: x.clone().Protocol().unwrap(),
                internal_client: x.clone().InternalClient().unwrap(),
                description: x.clone().Description().unwrap(),
            })
            .collect::<Vec<StaticPortMapping>>();
        let mut out: Box<dyn Write> = match std::env::args().nth(1) {
            Some(file_name) => Box::new(
                std::fs::OpenOptions::new()
                    .create(true)
                    .write(true)
                    .truncate(true)
                    .open(file_name)
                    .map_err(|_| Error::OK)?,
            ),
            None => Box::new(std::io::stdout()),
        };
        let _ = out.write_all(format!("{:#?}", enum_static_port_mapping).as_bytes());
        CoUninitialize();
    }
    Ok(())
}

#[derive(Debug)]
struct StaticPortMapping {
    external_port: i32,
    internal_port: i32,
    protocol: BSTR,
    internal_client: BSTR,
    description: BSTR,
}
