/**
 * WinToMyanmar3 Converter
 * Converts Win font encoding to Myanmar3 Unicode encoding
 * Ported from C# to TypeScript for Bun runtime
 */

// Font mapping entries as tuples [key, value]
// These are ordered intentionally - later entries override earlier ones
// Longer keys should be processed before shorter keys to avoid partial replacements
const fontMappingEntries: [string, string][] = [
  // Kinzi and special combinations (process first - longer combinations)
  ["ps", "\u1008"], // ဈ
  ["Bo", "\u1029"], // ဩ
  ["Mo", "\u1029"], // ဩ
  ["OD", "\u1026"], // ဦ
  ["ÍD", "\u1026"], // ဦ
  ["aBomf", "\u102A"], // ဪ
  ["aMomf", "\u102A"], // ဪ

  // Two-character combinations
  ["F", "\u1004\u103A\u1039"], // kinzi
  ["ø", "\u1036\u1004\u103A\u1039"],
  ["Ð", "\u1004\u103A\u1039\u102E"],
  ["Ø", "\u1004\u103A\u1039\u102D"],
  ["ð", "\u102D\u1036"],
  ["R", "\u103B\u103D"], // ျွ (ya + wa)
  ["Q", "\u103B\u103E"], // ျှ (ya + ha)
  ["W", "\u103B\u103D\u103E"], // ျွှ (ya + wa + ha)
  ["<", "\u103C\u103D"], // ြွ (ra + wa)
  [">", "\u103C\u103D"], // ြွ (ra + wa)
  ["ê", "\u103C\u102F"], // ြု (ra + u)
  ["û", "\u103C\u102F"], // ြု (ra + u)
  ["T", "\u103D\u103D\u103E"], // ွွှ
  ["I", "\u103E\u102F"], // ှု (ha + u)
  ["ª", "\u103E\u1030"], // ှူ (ha + uu)
  [":", "\u102B\u103A"], // ါ် (tall AA + asat)

  // Special characters
  ["þ", "\u1024"], // ဤ
  ["£", "\u1023"], // ဣ
  ["O", "\u1025"], // ဥ
  ["Í", "\u1025"], // ဥ
  ["Ó", "\u1009\u102C"], // ဉာ
  ["ó", "\u103F"], // ဿ
  ["@", "\u100F\u1039\u100D"], // ဏ္ဍ
  ["|", "\u100B\u1039\u100C"], // ဋ္ဌ
  ["¥", "\u100B\u1039\u100B"], // ဋ္ဋ
  ["×", "\u100D\u1039\u100D"], // ဍ္ဍ
  ["¹", "\u100E\u1039\u100D"], // ဍ္ဎ
  ["¿", "?"],
  ["µ", "!"],
  ["μ", "!"],
  ["«", "["],
  ["»", "]"],
  ["ç", ","],
  ["$", "\u1000\u1019\u1015\u103A"], // ကျပ် (kyat)
  ["_", "*"],
  ["ƒ", "\u1041\u2044\u1042"], // ၁/၂
  ["„", "\u1041\u2044\u1043"], // ၁/၃
  ["…", "\u1042\u2044\u1043"], // ၂/၃
  ["†", "\u1041\u2044\u1044"], // ၁/၄
  ["‡", "\u1043\u2044\u1044"], // ၃/၄
  ["ˆ", "\u1041\u2044\u1045"], // ၁/၅
  ["‰", "\u1042\u2044\u1045"], // ၂/၅
  ["Š", "\u1043\u2044\u1045"], // ၃/၅
  ["‹", "\u1044\u2044\u1045"], // ၄/၅
  ["ü", "\u104C"], // ၌
  ["í", "\u104D"], // ၍
  ["¤", "\u104E"], // ၎
  ["\\", "\u104F"], // ၏

  // Stacked consonants (virama combinations)
  ["ú", "\u1039\u1000"], // ္က
  ["©", "\u1039\u1001"], // ္ခ
  ["¾", "\u1039\u1002"], // ္ဂ
  ["¢", "\u1039\u1003"], // ္ဃ
  ["ö", "\u1039\u1005"], // ္စ
  ["ä", "\u1039\u1006"], // ္ဆ
  ["Æ", "\u1039\u1007"], // ္ဇ
  ["Ñ", "\u1039\u1008"], // ္ဈ
  ["³", "\u1039\u100C"], // ္ဌ
  ["²", "\u1039\u100D"], // ္ဍ
  ["Ü", "\u1039\u1015"], // ္ပ
  ["Ö", "\u1039\u100F"], // ္ဏ
  ["Å", "\u1039\u1010"], // ္တ
  ["å", "\u1039\u1010"], // ္တ
  ["¦", "\u1039\u1011"], // ္ထ
  ["¬", "\u1039\u1011"], // ္ထ
  ["´", "\u1039\u1012"], // ္ဒ
  ["¨", "\u1039\u1013"], // ္ဓ
  ["é", "\u1039\u1014"], // ္န
  ["æ", "\u1039\u1016"], // ္ဖ
  ["Ç", "\u1039\u1018"], // ္ဘ
  ["®", "\u1039\u1019"], // ္မ

  // Consonants
  ["u", "\u1000"], // က
  ["c", "\u1001"], // ခ
  ["*", "\u1002"], // ဂ
  ["C", "\u1003"], // ဃ
  ["i", "\u1004"], // င
  ["p", "\u1005"], // စ
  ["q", "\u1006"], // ဆ
  ["Z", "\u1007"], // ဇ
  ["Ú", "\u1009"], // ဉ
  ["n", "\u100A"], // ည
  ["ñ", "\u100A"], // ည
  ["#", "\u100B"], // ဋ
  ["X", "\u100C"], // ဌ
  ["!", "\u100D"], // ဍ
  ["¡", "\u100E"], // ဎ
  ["P", "\u100F"], // ဏ
  ["w", "\u1010"], // တ
  ["x", "\u1011"], // ထ
  ["'", "\u1012"], // ဒ (apostrophe)
  ['"', "\u1013"], // ဓ (quotation mark)
  ["e", "\u1014"], // န
  ["E", "\u1014"], // န
  ["y", "\u1015"], // ပ
  ["z", "\u1016"], // ဖ
  ["A", "\u1017"], // ဗ
  ["b", "\u1018"], // ဘ
  ["r", "\u1019"], // မ
  [",", "\u101A"], // ယ
  ["&", "\u101B"], // ရ
  ["½", "\u101B"], // ရ
  ["v", "\u101C"], // လ
  ["o", "\u101E"], // သ
  ["[", "\u101F"], // ဟ
  ["V", "\u1020"], // ဠ
  ["t", "\u1021"], // အ

  // Medial consonants
  ["s", "\u103B"], // ျ (ya)
  ["ß", "\u103B"], // ျ (ya)
  ["`", "\u103C"], // ြ (ra)
  ["j", "\u103C"], // ြ (ra)
  ["~", "\u103C"], // ြ (ra)
  ["B", "\u103C"], // ြ (ra)
  ["M", "\u103C"], // ြ (ra)
  ["N", "\u103C"], // ြ (ra)
  ["G", "\u103D"], // ွ (wa)
  ["S", "\u103E"], // ှ (ha)
  ["§", "\u103E"], // ှ (ha)

  // Independent vowels
  ["{", "\u1027"], // ဧ

  // Dependent vowels
  ["g", "\u102B"], // ါ (tall AA)
  ["m", "\u102C"], // ာ (AA)
  ["d", "\u102D"], // ိ (vi)
  ["D", "\u102E"], // ီ (ii)
  ["k", "\u102F"], // ု (u)
  ["K", "\u102F"], // ု (u)
  ["l", "\u1030"], // ူ (uu)
  ["L", "\u1030"], // ူ (uu)
  ["a", "\u1031"], // ေ (ve)
  ["J", "\u1032"], // ဲ (ai)
  ["H", "\u1036"], // ံ (ans)

  // Tone marks and signs
  ["f", "\u103A"], // ် (asat)
  ["Y", "\u1037"], // ့ (db)
  ["U", "\u1037"], // ့ (db)
  ["h", "\u1037"], // ့ (db)
  [";", "\u1038"], // း (visarga)

  // Digits (process last - overrides consonant mapping for "0")
  ["0", "\u1040"], // ၀
  ["1", "\u1041"], // ၁
  ["2", "\u1042"], // ၂
  ["3", "\u1043"], // ၃
  ["4", "\u1044"], // ၄
  ["5", "\u1045"], // ၅
  ["6", "\u1046"], // ၆
  ["7", "\u1047"], // ၇
  ["8", "\u1048"], // ၈
  ["9", "\u1049"], // ၉

  // Punctuation and special characters (process last)
  ["/", "\u104B"], // ။
  ["?", "\u104A"], // ၊
  ["]", "'"],
  ["}", "'"],
  ["^", "/"],
];

/**
 * Apply font mapping to convert Win encoding to intermediate Unicode
 */
function applyFontMapping(input: string): string {
  let result = input;
  // Process entries in order - later entries can override earlier ones
  for (const [winChar, myanmar3Char] of fontMappingEntries) {
    result = result.split(winChar).join(myanmar3Char);
  }
  return result;
}

/**
 * Remove duplicate diacritical marks
 */
function childdeldul(match: string): string {
  if (match.length > 1) {
    return match[0];
  }
  return match;
}

/**
 * Apply corrections to the converted text
 * Ported from correct.cs Correction1 function
 */
function correction1(input: string): string {
  const C = [
    "\u1000", "\u1001", "\u1002", "\u1003", "\u1004", "\u1005", "\u1006",
    "\u1007", "\u1008", "\u1009", "\u100A", "\u100B", "\u100C", "\u100D",
    "\u100E", "\u100F", "\u1010", "\u1011", "\u1012", "\u1013", "\u1014",
    "\u1015", "\u1016", "\u1017", "\u1018", "\u1019", "\u101A", "\u101B",
    "\u101C", "\u101D", "\u101E", "\u101F", "\u1020", "\u1021"
  ];
  const M = ["\u103B", "\u103C", "\u103D", "\u103E"];
  const V = [
    "\u102B", "\u102C", "\u102D", "\u102E", "\u102F", "\u1030", "\u1031",
    "\u1032", "\u1036"
  ];
  const IV = ["\u1023", "\u1024", "\u1025", "\u1026", "\u1027", "\u1029", "\u102A", "\u104E"];
  const T = ["\u1037", "\u1038", "\u103A", "\u1039"];
  const D = ["\u1040", "\u1041", "\u1042", "\u1043", "\u1044", "\u1045", "\u1046", "\u1047", "\u1048", "\u1049"];

  let unistr = input;

  // Remove duplicate diacritical marks
  const pattern = /\u102D+|\u102E+|\u103D+|\u103E+|\u1032+|\u1037+|\u1036+|\u103A+/g;
  unistr = unistr.replace(pattern, childdeldul);

  // Specific character combinations
  unistr = unistr.replace(new RegExp(C[5] + M[0], "g"), C[8]);
  unistr = unistr.replace(new RegExp(C[30] + M[1], "g"), IV[5]);
  unistr = unistr.replace(new RegExp(C[30] + M[1] + V[6] + V[1] + T[2], "g"), IV[6]);
  unistr = unistr.replace(new RegExp(IV[5] + V[6] + V[1] + T[2], "g"), IV[6]);
  unistr = unistr.replace(new RegExp(IV[2] + V[3], "g"), IV[3]);
  unistr = unistr.replace(new RegExp(IV[2] + T[3], "g"), C[9] + T[3]);
  unistr = unistr.replace(new RegExp(IV[2] + T[2], "g"), C[9] + T[2]);
  unistr = unistr.replace(new RegExp(IV[2] + V[1], "g"), C[9] + V[1]);
  unistr = unistr.replace(new RegExp(D[4] + C[4] + T[2] + T[1], "g"), IV[7] + C[4] + T[2] + T[1]);
  unistr = unistr.replace(new RegExp(T[0] + T[2], "g"), T[2] + T[0]);
  unistr = unistr.replace(new RegExp(T[1] + T[2], "g"), T[2] + T[1]);

  // Medial reordering
  unistr = unistr.replace(new RegExp(M[3] + M[0], "g"), M[0] + M[3]);
  unistr = unistr.replace(new RegExp(M[3] + M[1], "g"), M[1] + M[3]);
  unistr = unistr.replace(new RegExp(M[3] + M[2], "g"), M[2] + M[3]);
  unistr = unistr.replace(new RegExp(M[2] + M[0], "g"), M[0] + M[2]);
  unistr = unistr.replace(new RegExp(M[2] + M[1], "g"), M[1] + M[2]);
  unistr = unistr.replace(
    new RegExp(
      M[3] + M[2] + M[1] + "|" + M[3] + M[1] + M[2] + "|" +
      M[2] + M[3] + M[1] + "|" + M[2] + M[1] + M[3] + "|" +
      M[1] + M[3] + M[2], "g"
    ),
    M[1] + M[2] + M[3]
  );
  unistr = unistr.replace(
    new RegExp(
      M[3] + M[2] + M[0] + "|" + M[3] + M[0] + M[2] + "|" +
      M[2] + M[3] + M[0] + "|" + M[2] + M[0] + M[3] + "|" +
      M[0] + M[3] + M[2], "g"
    ),
    M[0] + M[2] + M[3]
  );

  // Vowel reordering
  unistr = unistr.replace(new RegExp(V[8] + V[4], "g"), V[4] + V[8]);
  unistr = unistr.replace(new RegExp(V[4] + V[2], "g"), V[2] + V[4]);
  unistr = unistr.replace(new RegExp(V[8] + V[2], "g"), V[2] + V[8]);
  unistr = unistr.replace(new RegExp(T[0] + V[4], "g"), V[4] + T[0]);
  unistr = unistr.replace(new RegExp(T[0] + V[7], "g"), V[7] + T[0]);
  unistr = unistr.replace(new RegExp(T[0] + V[8], "g"), V[8] + T[0]);

  // Contracted words
  unistr = unistr.replace(
    new RegExp(C[26] + V[6] + V[1] + C[0] + M[0] + T[2] + V[1], "g"),
    C[26] + V[6] + V[1] + C[0] + T[2] + M[0] + V[1]
  );
  unistr = unistr.replace(new RegExp(C[20] + V[4] + T[2], "g"), C[20] + T[2] + V[4]);

  // Remove double asat
  unistr = unistr.replace(/\u103A\u103A/g, "\u103A");

  // Recognition of digit as consonant
  unistr = unistr.replace(new RegExp(D[0] + T[2], "g"), C[29] + T[2]);
  unistr = unistr.replace(new RegExp(D[7] + T[2], "g"), C[27] + T[2]);
  unistr = unistr.replace(new RegExp(D[8] + T[2], "g"), C[2] + T[2]);

  unistr = unistr.replace(new RegExp(D[0] + T[3], "g"), C[29] + T[3]);
  unistr = unistr.replace(new RegExp(D[7] + T[3], "g"), C[27] + T[3]);
  unistr = unistr.replace(new RegExp(D[8] + T[3], "g"), C[2] + T[3]);

  // Digit + vowel combinations
  const vowelRange = V[0] + "-" + V[8];
  unistr = unistr.replace(new RegExp(D[0] + "(?<vowel>[" + vowelRange + "])", "g"), C[29] + "$<vowel>");
  unistr = unistr.replace(new RegExp(D[7] + "(?<vowel>[" + vowelRange + "])", "g"), C[27] + "$<vowel>");
  unistr = unistr.replace(new RegExp(D[8] + "(?<vowel>[" + vowelRange + "])", "g"), C[2] + "$<vowel>");

  // Digit + medial combinations
  const medialRange = M[0] + "-" + M[3];
  unistr = unistr.replace(new RegExp(D[0] + "(?<medial>[" + medialRange + "])", "g"), C[29] + "$<medial>");
  unistr = unistr.replace(new RegExp(D[7] + "(?<medial>[" + medialRange + "])", "g"), C[27] + "$<medial>");
  unistr = unistr.replace(new RegExp(D[8] + "(?<medial>[" + medialRange + "])", "g"), C[2] + "$<medial>");

  // Digit + final combinations
  unistr = unistr.replace(
    new RegExp(D[0] + "(?<finale>([\\u1000-\\u1031][\\u1039-\\u103A]))", "g"),
    C[29] + "$<finale>"
  );
  unistr = unistr.replace(
    new RegExp(D[7] + "(?<finale>([\\u1000-\\u1031][\\u1039-\\u103A]))", "g"),
    C[27] + "$<finale>"
  );
  unistr = unistr.replace(
    new RegExp(D[8] + "(?<finale>([\\u1000-\\u1031][\\u1039-\\u103A]))", "g"),
    C[2] + "$<finale>"
  );

  // Final reordering
  unistr = unistr.replace(
    /(?<upper>[\u102D\u102E\u1036\u1032])(?<M>[\u103B-\u103E]+)/g,
    "$<M>$<upper>"
  );
  unistr = unistr.replace(
    /(?<DVs>[\u1036\u1037\u1038]+)(?<lower>[\u102F\u1030])/g,
    "$<lower>$<DVs>"
  );
  unistr = unistr.replace("့်", "့်");

  return unistr;
}

/**
 * Main conversion function from Win encoding to Myanmar3 Unicode
 * @param input - The input string in Win font encoding
 * @returns The converted string in Myanmar3 Unicode
 */
export function winToMyanmar3(input: string): string {
  // Apply font mapping
  let unistr = applyFontMapping(input);

  // Reordering kinzi
  const conPattern = "[uc*CipqZnñÍÚ#X!¡Pwx\\'\\\"eE½yzAbr,&v0o[Vt|ó]";
  unistr = unistr.replace(
    new RegExp(`(?<E>a)?(?<R>j)?(?<con>${conPattern})\\u1004\\u103A\\u1039`, "g"),
    "\\u1004\\u103A\\u1039$<E>$<R>$<con>"
  );
  unistr = unistr.replace(
    new RegExp(`(?<E>a)?(?<R>j)?(?<con>${conPattern})Ø`, "g"),
    "F$<E>$<R>$<con>d"
  );
  unistr = unistr.replace(
    new RegExp(`(?<E>a)?(?<R>j)?(?<con>${conPattern})Ð`, "g"),
    "F$<E>$<R>$<con>D"
  );
  unistr = unistr.replace(
    new RegExp(`(?<E>a)?(?<R>j)?(?<con>${conPattern})ø`, "g"),
    "F$<E>$<R>$<con>H"
  );

  // Reordering Ra
  unistr = unistr.replace(
    /(?<R>\u103C)(?<Wa>\u103D)?(?<Ha>\u103E)?(?<U>\u102F)?(?<con>[က-အ])(?<scon>\u1039[က-အ])?/g,
    "$<con>$<scon>$<R>$<Wa>$<Ha>$<U>"
  );

  // Zero and wa handling
  // Convert \u1040 (Myanmar digit 0) to \u101D (ဝ) when adjacent to Myanmar characters
  // Convert \u1047 (Myanmar digit 7) to \u101B (ရ) when adjacent to Myanmar characters
  const mmPattern = "[\\u1000-\\u101C\\u101E-\\u102A\\u102C\\u102E-\\u103F\\u104C-\\u109F\\u0020]";
  unistr = unistr.replace(
    new RegExp(`(?<z>\\u1040)(?=(?<Mm>${mmPattern}))|(?<=(?<Mm>${mmPattern}))(?<z>\\u1040)`, "g"),
    "\u101D"
  );
  unistr = unistr.replace(
    new RegExp(`(?<z>\\u1047)(?=(?<Mm>${mmPattern}))|(?<=(?<Mm>${mmPattern}))(?<z>\\u1047)`, "g"),
    "\u101B"
  );

  // Final reordering for storage order
  unistr = unistr.replace(
    /(?<E>\u1031)?(?<con>[က-အ])(?<scon>\u1039[က-အ])?(?<upper>[\u102D\u102E\u1032\u1036])?(?<DVs>[\u1037\u1038]){0,2}(?<M>[\u103B-\u103E]*)(?<lower>[\u102F\u1030])?(?<upper2>[\u102D\u102E\u1032])?/g,
    "$<con>$<scon>$<M>$<E>$<upper>$<lower>$<DVs>"
  );

  // Apply corrections
  unistr = correction1(unistr);

  return unistr;
}

// Export for Bun/Esm
export default { winToMyanmar3 };
