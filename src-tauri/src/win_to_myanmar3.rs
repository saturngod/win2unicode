use regex::Regex;

const FONT_MAPPING_ENTRIES: &[(&str, &str)] = &[
    // Kinzi and special combinations (process first - longer combinations)
    ("ps", "\u{1008}"),
    ("Bo", "\u{1029}"),
    ("Mo", "\u{1029}"),
    ("OD", "\u{1026}"),
    ("\u{00CD}D", "\u{1026}"),
    ("aBomf", "\u{102A}"),
    ("aMomf", "\u{102A}"),

    // Two-character combinations
    ("F", "\u{1004}\u{103A}\u{1039}"),
    ("\u{00F8}", "\u{1036}\u{1004}\u{103A}\u{1039}"),
    ("\u{00D0}", "\u{1004}\u{103A}\u{1039}\u{102E}"),
    ("\u{00D8}", "\u{1004}\u{103A}\u{1039}\u{102D}"),
    ("\u{00F0}", "\u{102D}\u{1036}"),
    ("R", "\u{103B}\u{103D}"),
    ("Q", "\u{103B}\u{103E}"),
    ("W", "\u{103B}\u{103D}\u{103E}"),
    ("<", "\u{103C}\u{103D}"),
    (">", "\u{103C}\u{103D}"),
    ("\u{00EA}", "\u{103C}\u{102F}"),
    ("\u{00FB}", "\u{103C}\u{102F}"),
    ("Bu", "\u{1000}\u{103C}"),
    ("T", "\u{103D}\u{103D}\u{103E}"),
    ("I", "\u{103E}\u{102F}"),
    ("\u{00AA}", "\u{103E}\u{1030}"),
    (":", "\u{102B}\u{103A}"),

    // Special characters
    ("\u{00FE}", "\u{1024}"),
    ("\u{00A3}", "\u{1023}"),
    ("O", "\u{1025}"),
    ("\u{00CD}", "\u{1025}"),
    ("\u{00D3}", "\u{1009}\u{102C}"),
    ("\u{00F3}", "\u{103F}"),
    ("@", "\u{100F}\u{1039}\u{100D}"),
    ("|", "\u{100B}\u{1039}\u{100C}"),
    ("\u{00A5}", "\u{100B}\u{1039}\u{100B}"),
    ("\u{00D7}", "\u{100D}\u{1039}\u{100D}"),
    ("\u{00B9}", "\u{100E}\u{1039}\u{100D}"),
    ("\u{00BF}", "?"),
    ("\u{00B5}", "!"),
    ("\u{03BC}", "!"),
    ("$", "\u{1000}\u{1019}\u{1015}\u{103A}"),
    ("_", "*"),
    ("\u{0192}", "\u{1041}\u{2044}\u{1042}"),
    ("\u{201E}", "\u{1041}\u{2044}\u{1043}"),
    ("\u{2026}", "\u{1042}\u{2044}\u{1043}"),
    ("\u{2020}", "\u{1041}\u{2044}\u{1044}"),
    ("\u{2021}", "\u{1043}\u{2044}\u{1044}"),
    ("\u{02C6}", "\u{1041}\u{2044}\u{1045}"),
    ("\u{2030}", "\u{1042}\u{2044}\u{1045}"),
    ("\u{0160}", "\u{1043}\u{2044}\u{1045}"),
    ("\u{2039}", "\u{1044}\u{2044}\u{1045}"),
    ("\u{00FC}", "\u{104C}"),
    ("\u{00ED}", "\u{104D}"),
    ("\u{00A4}", "\u{104E}"),
    ("\\", "\u{104F}"),

    // Stacked consonants (virama combinations)
    ("\u{00FA}", "\u{1039}\u{1000}"),
    ("\u{00A9}", "\u{1039}\u{1001}"),
    ("\u{00BE}", "\u{1039}\u{1002}"),
    ("\u{00A2}", "\u{1039}\u{1003}"),
    ("\u{00F6}", "\u{1039}\u{1005}"),
    ("\u{00E4}", "\u{1039}\u{1006}"),
    ("\u{00C6}", "\u{1039}\u{1007}"),
    ("\u{00D1}", "\u{1039}\u{1008}"),
    ("\u{00B3}", "\u{1039}\u{100C}"),
    ("\u{00B2}", "\u{1039}\u{100D}"),
    ("\u{00DC}", "\u{1039}\u{1015}"),
    ("\u{00D6}", "\u{1039}\u{100F}"),
    ("\u{00C5}", "\u{1039}\u{1010}"),
    ("\u{00E5}", "\u{1039}\u{1010}"),
    ("\u{00A6}", "\u{1039}\u{1011}"),
    ("\u{00AC}", "\u{1039}\u{1011}"),
    ("\u{00B4}", "\u{1039}\u{1012}"),
    ("\u{00A8}", "\u{1039}\u{1013}"),
    ("\u{00E9}", "\u{1039}\u{1014}"),
    ("\u{00E6}", "\u{1039}\u{1016}"),
    ("\u{00C7}", "\u{1039}\u{1018}"),
    ("\u{00AE}", "\u{1039}\u{1019}"),

    // Consonants
    ("u", "\u{1000}"),
    ("c", "\u{1001}"),
    ("*", "\u{1002}"),
    ("C", "\u{1003}"),
    ("i", "\u{1004}"),
    ("p", "\u{1005}"),
    ("q", "\u{1006}"),
    ("Z", "\u{1007}"),
    ("\u{00DA}", "\u{1009}"),
    ("n", "\u{100A}"),
    ("\u{00F1}", "\u{100A}"),
    ("#", "\u{100B}"),
    ("X", "\u{100C}"),
    ("!", "\u{100D}"),
    ("\u{00A1}", "\u{100E}"),
    ("P", "\u{100F}"),
    ("w", "\u{1010}"),
    ("x", "\u{1011}"),
    ("'", "\u{1012}"),
    ("\"", "\u{1013}"),
    ("e", "\u{1014}"),
    ("E", "\u{1014}"),
    ("y", "\u{1015}"),
    ("z", "\u{1016}"),
    ("A", "\u{1017}"),
    ("b", "\u{1018}"),
    ("r", "\u{1019}"),
    (",", "\u{101A}"),
    ("&", "\u{101B}"),
    ("\u{00BD}", "\u{101B}"),
    ("v", "\u{101C}"),
    ("o", "\u{101E}"),
    ("[", "\u{101F}"),
    ("V", "\u{1020}"),
    ("t", "\u{1021}"),

    // Medial consonants
    ("s", "\u{103B}"),
    ("\u{00DF}", "\u{103B}"),
    ("`", "\u{103C}"),
    ("j", "\u{103C}"),
    ("~", "\u{103C}"),
    ("B", "\u{103C}"),
    ("M", "\u{103C}"),
    ("N", "\u{103C}"),
    ("G", "\u{103D}"),
    ("S", "\u{103E}"),
    ("\u{00A7}", "\u{103E}"),

    // Independent vowels
    ("{", "\u{1027}"),

    // Dependent vowels
    ("g", "\u{102B}"),
    ("m", "\u{102C}"),
    ("d", "\u{102D}"),
    ("D", "\u{102E}"),
    ("k", "\u{102F}"),
    ("K", "\u{102F}"),
    ("l", "\u{1030}"),
    ("L", "\u{1030}"),
    ("a", "\u{1031}"),
    ("J", "\u{1032}"),
    ("H", "\u{1036}"),

    // Tone marks and signs
    ("f", "\u{103A}"),
    ("Y", "\u{1037}"),
    ("U", "\u{1037}"),
    ("h", "\u{1037}"),
    (";", "\u{1038}"),

    // Digits (process last - overrides consonant mapping for "0")
    ("0", "\u{1040}"),
    ("1", "\u{1041}"),
    ("2", "\u{1042}"),
    ("3", "\u{1043}"),
    ("4", "\u{1044}"),
    ("5", "\u{1045}"),
    ("6", "\u{1046}"),
    ("7", "\u{1047}"),
    ("8", "\u{1048}"),
    ("9", "\u{1049}"),

    // Punctuation and special characters (process last)
    ("/", "\u{104B}"),
    ("?", "\u{104A}"),
    ("]", "'"),
    ("}", "'"),
    ("^", "/"),
];

fn apply_font_mapping(input: &str) -> String {
    let mut result = input.to_string();
    for (win_char, myanmar_char) in FONT_MAPPING_ENTRIES {
        result = result.replace(win_char, myanmar_char);
    }
    result
}

fn childdeldul(match_str: &str) -> &str {
    if match_str.len() > 1 {
        &match_str[0..match_str.chars().next().unwrap().len_utf8()]
    } else {
        match_str
    }
}

fn correction1(input: &str) -> String {
    let c = [
        "\u{1000}", "\u{1001}", "\u{1002}", "\u{1003}", "\u{1004}", "\u{1005}", "\u{1006}",
        "\u{1007}", "\u{1008}", "\u{1009}", "\u{100A}", "\u{100B}", "\u{100C}", "\u{100D}",
        "\u{100E}", "\u{100F}", "\u{1010}", "\u{1011}", "\u{1012}", "\u{1013}", "\u{1014}",
        "\u{1015}", "\u{1016}", "\u{1017}", "\u{1018}", "\u{1019}", "\u{101A}", "\u{101B}",
        "\u{101C}", "\u{101D}", "\u{101E}", "\u{101F}", "\u{1020}", "\u{1021}",
    ];
    let m = ["\u{103B}", "\u{103C}", "\u{103D}", "\u{103E}"];
    let v = [
        "\u{102B}", "\u{102C}", "\u{102D}", "\u{102E}", "\u{102F}", "\u{1030}", "\u{1031}",
        "\u{1032}", "\u{1036}",
    ];
    let iv = ["\u{1023}", "\u{1024}", "\u{1025}", "\u{1026}", "\u{1027}", "\u{1029}", "\u{102A}", "\u{104E}"];
    let t = ["\u{1037}", "\u{1038}", "\u{103A}", "\u{1039}"];
    let d = ["\u{1040}", "\u{1041}", "\u{1042}", "\u{1043}", "\u{1044}", "\u{1045}", "\u{1046}", "\u{1047}", "\u{1048}", "\u{1049}"];

    let mut unistr = input.to_string();

    // Remove duplicate diacritical marks
    let dup_re = Regex::new("\u{102D}+|\u{102E}+|\u{103D}+|\u{103E}+|\u{1032}+|\u{1037}+|\u{1036}+|\u{103A}+").unwrap();
    unistr = dup_re
        .replace_all(&unistr, |caps: &regex::Captures| {
            childdeldul(caps.get(0).unwrap().as_str()).to_string()
        })
        .to_string();

    // Specific character combinations
    unistr = unistr.replace(&format!("{}{}", c[5], m[0]), c[8]);
    unistr = unistr.replace(&format!("{}{}", c[30], m[1]), iv[5]);
    unistr = unistr.replace(&format!("{}{}{}{}{}", c[30], m[1], v[6], v[1], t[2]), iv[6]);
    unistr = unistr.replace(&format!("{}{}{}{}", iv[5], v[6], v[1], t[2]), iv[6]);
    unistr = unistr.replace(&format!("{}{}", iv[2], v[3]), iv[3]);
    unistr = unistr.replace(&format!("{}{}", iv[2], t[3]), &format!("{}{}", c[9], t[3]));
    unistr = unistr.replace(&format!("{}{}", iv[2], t[2]), &format!("{}{}", c[9], t[2]));
    unistr = unistr.replace(&format!("{}{}", iv[2], v[1]), &format!("{}{}", c[9], v[1]));
    unistr = unistr.replace(&format!("{}{}{}{}", d[4], c[4], t[2], t[1]), &format!("{}{}{}{}", iv[7], c[4], t[2], t[1]));
    unistr = unistr.replace(&format!("{}{}", t[0], t[2]), &format!("{}{}", t[2], t[0]));
    unistr = unistr.replace(&format!("{}{}", t[1], t[2]), &format!("{}{}", t[2], t[1]));

    // Medial reordering
    unistr = unistr.replace(&format!("{}{}", m[3], m[0]), &format!("{}{}", m[0], m[3]));
    unistr = unistr.replace(&format!("{}{}", m[3], m[1]), &format!("{}{}", m[1], m[3]));
    unistr = unistr.replace(&format!("{}{}", m[3], m[2]), &format!("{}{}", m[2], m[3]));
    unistr = unistr.replace(&format!("{}{}", m[2], m[0]), &format!("{}{}", m[0], m[2]));
    unistr = unistr.replace(&format!("{}{}", m[2], m[1]), &format!("{}{}", m[1], m[2]));
    unistr = Regex::new(&format!(
        "{}{}{}|{}{}{}|{}{}{}|{}{}{}|{}{}{}",
        m[3], m[2], m[1],
        m[3], m[1], m[2],
        m[2], m[3], m[1],
        m[2], m[1], m[3],
        m[1], m[3], m[2],
    ))
    .unwrap()
    .replace_all(&unistr, &format!("{}{}{}", m[1], m[2], m[3]))
    .to_string();
    unistr = Regex::new(&format!(
        "{}{}{}|{}{}{}|{}{}{}|{}{}{}|{}{}{}",
        m[3], m[2], m[0],
        m[3], m[0], m[2],
        m[2], m[3], m[0],
        m[2], m[0], m[3],
        m[0], m[3], m[2],
    ))
    .unwrap()
    .replace_all(&unistr, &format!("{}{}{}", m[0], m[2], m[3]))
    .to_string();

    // Vowel reordering
    unistr = unistr.replace(&format!("{}{}", v[8], v[4]), &format!("{}{}", v[4], v[8]));
    unistr = unistr.replace(&format!("{}{}", v[4], v[2]), &format!("{}{}", v[2], v[4]));
    unistr = unistr.replace(&format!("{}{}", v[8], v[2]), &format!("{}{}", v[2], v[8]));
    unistr = unistr.replace(&format!("{}{}", t[0], v[4]), &format!("{}{}", v[4], t[0]));
    unistr = unistr.replace(&format!("{}{}", t[0], v[7]), &format!("{}{}", v[7], t[0]));
    unistr = unistr.replace(&format!("{}{}", t[0], v[8]), &format!("{}{}", v[8], t[0]));

    // Contracted words
    unistr = unistr.replace(
        &format!("{}{}{}{}{}{}{}", c[26], v[6], v[1], c[0], m[0], t[2], v[1]),
        &format!("{}{}{}{}{}{}{}", c[26], v[6], v[1], c[0], t[2], m[0], v[1]),
    );
    unistr = unistr.replace(&format!("{}{}{}", c[20], v[4], t[2]), &format!("{}{}{}", c[20], t[2], v[4]));

    // Remove double asat
    unistr = unistr.replace(&format!("{}{}", t[2], t[2]), t[2]);

    // Recognition of digit as consonant
    unistr = unistr.replace(&format!("{}{}", d[0], t[2]), &format!("{}{}", c[29], t[2]));
    unistr = unistr.replace(&format!("{}{}", d[7], t[2]), &format!("{}{}", c[27], t[2]));
    unistr = unistr.replace(&format!("{}{}", d[8], t[2]), &format!("{}{}", c[2], t[2]));

    unistr = unistr.replace(&format!("{}{}", d[0], t[3]), &format!("{}{}", c[29], t[3]));
    unistr = unistr.replace(&format!("{}{}", d[7], t[3]), &format!("{}{}", c[27], t[3]));
    unistr = unistr.replace(&format!("{}{}", d[8], t[3]), &format!("{}{}", c[2], t[3]));

    // Digit + vowel combinations
    let vowel_range = format!("{}-{}", v[0], v[8]);
    unistr = Regex::new(&format!("{}(?P<vowel>[{}])", d[0], vowel_range))
        .unwrap()
        .replace_all(&unistr, &format!("{}$vowel", c[29]))
        .to_string();
    unistr = Regex::new(&format!("{}(?P<vowel>[{}])", d[7], vowel_range))
        .unwrap()
        .replace_all(&unistr, &format!("{}$vowel", c[27]))
        .to_string();
    unistr = Regex::new(&format!("{}(?P<vowel>[{}])", d[8], vowel_range))
        .unwrap()
        .replace_all(&unistr, &format!("{}$vowel", c[2]))
        .to_string();

    // Digit + medial combinations
    let medial_range = format!("{}-{}", m[0], m[3]);
    unistr = Regex::new(&format!("{}(?P<medial>[{}])", d[0], medial_range))
        .unwrap()
        .replace_all(&unistr, &format!("{}$medial", c[29]))
        .to_string();
    unistr = Regex::new(&format!("{}(?P<medial>[{}])", d[7], medial_range))
        .unwrap()
        .replace_all(&unistr, &format!("{}$medial", c[27]))
        .to_string();
    unistr = Regex::new(&format!("{}(?P<medial>[{}])", d[8], medial_range))
        .unwrap()
        .replace_all(&unistr, &format!("{}$medial", c[2]))
        .to_string();

    // Digit + final combinations
    unistr = Regex::new(&format!("{}(?P<finale>([{}-{}][{}-{}]))", d[0], "\u{1000}", "\u{1031}", "\u{1039}", "\u{103A}"))
        .unwrap()
        .replace_all(&unistr, &format!("{}$finale", c[29]))
        .to_string();
    unistr = Regex::new(&format!("{}(?P<finale>([{}-{}][{}-{}]))", d[7], "\u{1000}", "\u{1031}", "\u{1039}", "\u{103A}"))
        .unwrap()
        .replace_all(&unistr, &format!("{}$finale", c[27]))
        .to_string();
    unistr = Regex::new(&format!("{}(?P<finale>([{}-{}][{}-{}]))", d[8], "\u{1000}", "\u{1031}", "\u{1039}", "\u{103A}"))
        .unwrap()
        .replace_all(&unistr, &format!("{}$finale", c[2]))
        .to_string();

    // Final reordering
    unistr = Regex::new("(?P<upper>[\u{102D}\u{102E}\u{1036}\u{1032}])(?P<M>[\u{103B}-\u{103E}]+)")
        .unwrap()
        .replace_all(&unistr, "$M$upper")
        .to_string();
    unistr = Regex::new("(?P<DVs>[\u{1036}\u{1037}\u{1038}]+)(?P<lower>[\u{102F}\u{1030}])")
        .unwrap()
        .replace_all(&unistr, "$lower$DVs")
        .to_string();

    // Original JS: unistr = unistr.replace("့်", "့်");
    unistr = unistr.replace("့်", "့်");

    unistr
}

fn is_mm_context(ch: char) -> bool {
    let code = ch as u32;
    (0x1000..=0x101C).contains(&code)
        || (0x101E..=0x102A).contains(&code)
        || code == 0x102C
        || (0x102E..=0x103F).contains(&code)
        || (0x104C..=0x109F).contains(&code)
        || code == 0x0020
}

fn replace_zero_wa(input: &str) -> String {
    let chars: Vec<char> = input.chars().collect();
    let mut out = String::with_capacity(chars.len());

    for i in 0..chars.len() {
        let ch = chars[i];
        let prev_mm = if i > 0 { is_mm_context(chars[i - 1]) } else { false };
        let next_mm = if i + 1 < chars.len() { is_mm_context(chars[i + 1]) } else { false };

        if ch == '\u{1040}' {
            if prev_mm || next_mm {
                out.push('\u{101D}');
            } else {
                out.push(ch);
            }
        } else if ch == '\u{1047}' {
            if prev_mm || next_mm {
                out.push('\u{101B}');
            } else {
                out.push(ch);
            }
        } else {
            out.push(ch);
        }
    }

    out
}

fn cleanup_before_convert(input: &str) -> String {
    let mut result = input.to_string();
    let replacements = [
        (" f", "f"),
        (" m", "m"),
        ("  ;", ";"),
        ("a ", "a"),
        (" D", "D"),
        (" d", "d"),
        (" F", "F"),
        (" S", "S"),
    ];
    for (from, to) in replacements {
        result = result.replace(from, to);
    }
    result
}

fn cleanup_after_convert(input: &str) -> String {
    let mut result = input.to_string();
    let replacements = [("«", "["), ("»", "]"), ("ç", ",")];
    for (from, to) in replacements {
        result = result.replace(from, to);
    }
    result
}

pub fn win_to_myanmar3(input: &str) -> String {
    // Cleanup before font mapping
    let cleaned = cleanup_before_convert(input);

    // Apply font mapping
    let mut unistr = apply_font_mapping(&cleaned);

    // Reordering kinzi
    let con_pattern = "(?:u|c|\\*|C|i|p|q|Z|n|\u{00F1}|\u{00CD}|\u{00DA}|#|X|!|\u{00A1}|P|w|x|'|\"|e|E|\u{00BD}|y|z|A|b|r|,|&|v|0|o|\\[|V|t|\\||\u{00F3})";
    let re = Regex::new(&format!("(?P<E>a)?(?P<R>j)?(?P<con>{})\u{1004}\u{103A}\u{1039}", con_pattern)).unwrap();
    unistr = re
        .replace_all(&unistr, "\u{1004}\u{103A}\u{1039}${E}${R}${con}")
        .to_string();
    let re = Regex::new(&format!("(?P<E>a)?(?P<R>j)?(?P<con>{})\u{00D8}", con_pattern)).unwrap();
    unistr = re.replace_all(&unistr, "F${E}${R}${con}d").to_string();
    let re = Regex::new(&format!("(?P<E>a)?(?P<R>j)?(?P<con>{})\u{00D0}", con_pattern)).unwrap();
    unistr = re.replace_all(&unistr, "F${E}${R}${con}D").to_string();
    let re = Regex::new(&format!("(?P<E>a)?(?P<R>j)?(?P<con>{})\u{00F8}", con_pattern)).unwrap();
    unistr = re.replace_all(&unistr, "F${E}${R}${con}H").to_string();

    // Reordering Ra
    let re = Regex::new("(?P<R>\u{103C})(?P<Wa>\u{103D})?(?P<Ha>\u{103E})?(?P<U>\u{102F})?(?P<con>[\u{1000}-\u{1021}])(?P<scon>\u{1039}[\u{1000}-\u{1021}])?")
        .unwrap();
    unistr = re
        .replace_all(&unistr, "${con}${scon}${R}${Wa}${Ha}${U}")
        .to_string();

    // Zero and wa handling
    unistr = replace_zero_wa(&unistr);

    // Final reordering for storage order
    let re = Regex::new("(?P<E>\u{1031})?(?P<con>[\u{1000}-\u{1021}])(?P<scon>\u{1039}[\u{1000}-\u{1021}])?(?P<upper>[\u{102D}\u{102E}\u{1032}\u{1036}])?(?P<DVs>[\u{1037}\u{1038}]){0,2}(?P<M>[\u{103B}-\u{103E}]*)(?P<lower>[\u{102F}\u{1030}])?(?P<upper2>[\u{102D}\u{102E}\u{1032}])?")
        .unwrap();
    unistr = re
        .replace_all(&unistr, "${con}${scon}${M}${E}${upper}${lower}${DVs}${upper2}")
        .to_string();

    // Apply corrections
    unistr = correction1(&unistr);

    // Cleanup after conversion (avoid conflicts during Win Innwa mapping)
    cleanup_after_convert(&unistr)
}
