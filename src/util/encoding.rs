use encoding_rs::{UTF_16BE, UTF_16LE};

fn try_repair_utf16be(s: &String) -> Option<String> {
    let looks_utf16be = s.as_bytes().chunks(2).all(|x| x.len() == 2 && x[0] == 0u8);
    // if every first byte is null, then this is a UTF16BE inside of a UTF8 string.
    if looks_utf16be {
        let (res, _, had_errors) = UTF_16BE.decode(s.as_bytes());
        if had_errors {
            return None;
        } else {
            return Some(res.into_owned());
        }
    } else {
        return None;
    }
}

fn try_repair_utf16le(s: &String) -> Option<String> {
    let looks_utf16le = s.as_bytes().chunks(2).all(|x| x.len() == 2 && x[1] == 0u8);

    // if every second byte is null, then this is a UTF16LE inside of a UTF8 string.
    if looks_utf16le {
        let (res, _, had_errors) = UTF_16LE.decode(s.as_bytes());
        if had_errors {
            return None;
        } else {
            return Some(res.into_owned());
        }
    } else {
        return None;
    }
}

/**
 * Calamine doesn't handle UTF16 encoded xls files well, and fails to decode
 * them. So, we instead get Rust Strings (which, in Rust, are UTF8) that
 * actually contain UTF16LE encoded data. This function attempts to detect and
 * repair those strings.
 */
pub fn repair_bad_encodings(s: String) -> String {
    try_repair_utf16be(&s)
        .or_else(|| try_repair_utf16le(&s))
        .unwrap_or(s)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_try_repair_utf16le_given_utf16le() {
        let bytes = vec![0x61, 0x00, 0x62, 0x00, 0x63, 0x00, 0x64, 0x00]; // 'abcd' in UTF16LE
        let s = String::from_utf8(bytes).unwrap();
        assert_eq!(try_repair_utf16le(&s), Some(String::from("abcd")));
    }

    #[test]
    fn test_try_repair_utf16le_given_utf16be() {
        let bytes = vec![0x00, 0x61, 0x00, 0x62, 0x00, 0x63, 0x00, 0x64]; // 'abcd' in UTF16BE
        let s = String::from_utf8(bytes).unwrap();
        assert_eq!(try_repair_utf16le(&s), None);
    }

    #[test]
    fn test_try_repair_utf16le_given_utf8() {
        let s = String::from("abcd"); // 'abcd' in UTF8
        assert_eq!(try_repair_utf16le(&s), None);
    }

    #[test]
    fn test_try_repair_utf16be_given_utf16be() {
        let bytes = vec![0x00, 0x61, 0x00, 0x62, 0x00, 0x63, 0x00, 0x64]; // 'abcd' in UTF16BE
        let s = String::from_utf8(bytes).unwrap();
        assert_eq!(try_repair_utf16be(&s), Some(String::from("abcd")));
    }

    #[test]
    fn test_try_repair_utf16be_given_utf16le() {
        let bytes = vec![0x61, 0x00, 0x62, 0x00, 0x63, 0x00, 0x64, 0x00]; // 'abcd' in UTF16LE
        let s = String::from_utf8(bytes).unwrap();
        assert_eq!(try_repair_utf16be(&s), None);
    }

    #[test]
    fn test_try_repair_utf16be_given_utf8() {
        let s = String::from("abcd"); // 'abcd' in UTF8
        assert_eq!(try_repair_utf16be(&s), None);
    }
}
