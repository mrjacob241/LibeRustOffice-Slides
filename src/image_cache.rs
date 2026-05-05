use std::{
    fs::{self, OpenOptions},
    io::{self, Write},
    path::{Path, PathBuf},
};

const MAGIC: &[u8; 8] = b"LROIC01\0";
const MAX_KEY_LEN: usize = 16 * 1024;
const MAX_MEDIA_TYPE_LEN: usize = 512;
const MAX_IMAGE_BYTES: usize = 128 * 1024 * 1024;

#[derive(Clone, Debug)]
pub struct CachedImage {
    pub media_type: String,
    pub bytes: Vec<u8>,
}

pub fn store_image(key: impl AsRef<str>, media_type: impl AsRef<str>, bytes: &[u8]) {
    if key.as_ref().is_empty() || bytes.is_empty() || bytes.len() > MAX_IMAGE_BYTES {
        return;
    }
    if let Err(error) = append_record(key.as_ref(), media_type.as_ref(), bytes) {
        eprintln!("image cache write failed: {error}");
    }
}

pub fn store_path(path: impl AsRef<Path>) {
    let path = path.as_ref();
    let Ok(bytes) = fs::read(path) else {
        return;
    };
    let key = path.to_string_lossy();
    let media_type = media_type_for_path(path);
    store_image(key.as_ref(), media_type, &bytes);
}

pub fn load_latest(key: impl AsRef<str>) -> Option<CachedImage> {
    let key = key.as_ref();
    if key.is_empty() {
        return None;
    }
    let bytes = fs::read(cache_path()).ok()?;
    let mut cursor = 0usize;
    let mut latest = None;

    while cursor < bytes.len() {
        let record_start = cursor;
        let magic = bytes.get(cursor..cursor + MAGIC.len())?;
        cursor += MAGIC.len();
        if magic != MAGIC {
            cursor = record_start + 1;
            continue;
        }

        let key_len = read_u32(&bytes, &mut cursor)? as usize;
        let media_type_len = read_u32(&bytes, &mut cursor)? as usize;
        let image_len = read_u64(&bytes, &mut cursor)? as usize;
        if key_len > MAX_KEY_LEN
            || media_type_len > MAX_MEDIA_TYPE_LEN
            || image_len > MAX_IMAGE_BYTES
        {
            return latest;
        }

        let key_bytes = take(&bytes, &mut cursor, key_len)?;
        let media_type_bytes = take(&bytes, &mut cursor, media_type_len)?;
        let image_bytes = take(&bytes, &mut cursor, image_len)?;

        if key_bytes == key.as_bytes() {
            latest = Some(CachedImage {
                media_type: String::from_utf8_lossy(media_type_bytes).into_owned(),
                bytes: image_bytes.to_vec(),
            });
        }
    }

    latest
}

pub fn media_type_for_path(path: impl AsRef<Path>) -> &'static str {
    match path
        .as_ref()
        .extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.to_ascii_lowercase())
        .as_deref()
    {
        Some("jpg") | Some("jpeg") => "image/jpeg",
        Some("gif") => "image/gif",
        Some("bmp") => "image/bmp",
        Some("webp") => "image/webp",
        Some("svg") => "image/svg+xml",
        _ => "image/png",
    }
}

pub fn extension_for_media_type(media_type: &str) -> &'static str {
    match media_type {
        "image/jpeg" => "jpg",
        "image/gif" => "gif",
        "image/bmp" => "bmp",
        "image/webp" => "webp",
        "image/svg+xml" => "svg",
        _ => "png",
    }
}

fn append_record(key: &str, media_type: &str, bytes: &[u8]) -> io::Result<()> {
    if let Some(parent) = cache_path().parent() {
        fs::create_dir_all(parent)?;
    }
    let mut file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(cache_path())?;
    file.write_all(MAGIC)?;
    file.write_all(&(key.len() as u32).to_le_bytes())?;
    file.write_all(&(media_type.len() as u32).to_le_bytes())?;
    file.write_all(&(bytes.len() as u64).to_le_bytes())?;
    file.write_all(key.as_bytes())?;
    file.write_all(media_type.as_bytes())?;
    file.write_all(bytes)?;
    Ok(())
}

fn cache_path() -> PathBuf {
    std::env::temp_dir().join("liberustoffice_slides_image_cache.bin")
}

fn read_u32(bytes: &[u8], cursor: &mut usize) -> Option<u32> {
    let data = take(bytes, cursor, 4)?;
    Some(u32::from_le_bytes(data.try_into().ok()?))
}

fn read_u64(bytes: &[u8], cursor: &mut usize) -> Option<u64> {
    let data = take(bytes, cursor, 8)?;
    Some(u64::from_le_bytes(data.try_into().ok()?))
}

fn take<'a>(bytes: &'a [u8], cursor: &mut usize, len: usize) -> Option<&'a [u8]> {
    let end = cursor.checked_add(len)?;
    let data = bytes.get(*cursor..end)?;
    *cursor = end;
    Some(data)
}
