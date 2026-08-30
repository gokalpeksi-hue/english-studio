let cards = [];
let currentIndex = 0;
let showingTurkish = false;
let femaleVoice = null;
let maleVoice = null;
let autoSpeak = false; // otomatik seslendirme kapalı (sadece 🔊 butonuyla okur)

// ——— Önbellekler ———
const dictCache = {};   // İngilizce → {main, groups, alts} Türkçe sözlük karşılıkları
const defCache = {};    // İngilizce → sözlük tanımları
const phraseCache = {}; // İngilizce cümle → Türkçe çeviri (örnek cümleler için)

// ——— Başarılı çevirileri kalıcı hafızada tut ———
// Bir kez çevrilen kelime/cümle, uygulama yeniden açıldığında da internetsiz gösterilir.
const TR_CACHE_KEY = "kelimeKarti_trCache_v1";
const TR_CACHE_MAX = 1500; // kayıt sınırı (localStorage şişmesin)

(function loadTrCache() {
    try {
        const saved = JSON.parse(localStorage.getItem(TR_CACHE_KEY));
        if (saved && typeof saved === "object") {
            if (saved.phrases) Object.assign(phraseCache, saved.phrases);
            if (saved.dict) {
                for (const k in saved.dict) {
                    dictCache[k] = { data: saved.dict[k] };
                }
            }
        }
    } catch (_) {}
})();

let trCacheSaveTimer = null;
function saveTrCache() {
    clearTimeout(trCacheSaveTimer);
    trCacheSaveTimer = setTimeout(() => {
        try {
            const phrases = {};
            const dict = {};
            let n = 0;
            for (const k in phraseCache) {
                if (typeof phraseCache[k] === "string" && n < TR_CACHE_MAX) { phrases[k] = phraseCache[k]; n++; }
            }
            for (const k in dictCache) {
                if (dictCache[k] && dictCache[k].data && n < TR_CACHE_MAX) { dict[k] = dictCache[k].data; n++; }
            }
            localStorage.setItem(TR_CACHE_KEY, JSON.stringify({ phrases, dict }));
        } catch (_) {} // depo doluysa sessizce vazgeç (bellek içi önbellek çalışmaya devam eder)
    }, 800);
}

// ——— Zamanlayıcılar ———
let hideTooltipTimer = null;
let hoverTimer = null;

// ——— Açık olan açıklama balonunun kaynağı (aç/kapa için) ———
let activeTooltipEl = null;

// ——— Örnek Cümle Modu ———
let isShowingExample = false;
let originalCardIndex = 0;

// ——— Seçim Modu (cümlede kelime grubu seçme) ———
let selectMode = false;    // 🖍 butonuyla açılıp kapanır
let selAnchorIdx = null;   // seçimin ilk dokunulan kelimesinin sırası

// =========================================================
// VERİ YÜKLEME
// =========================================================
// Yüklenen Excel/CSV'yi tarayıcı hafızasında kalıcı tutmak için anahtarlar
const STORAGE_KEY = "kelimeKarti_cards";
const STORAGE_NAME_KEY = "kelimeKarti_fileName";
const STORAGE_INDEX_KEY = "kelimeKarti_index";

// Kaldığımız kartın sırasını hafızaya kaydet
function saveCurrentIndex() {
    try {
        localStorage.setItem(STORAGE_INDEX_KEY, String(currentIndex));
    } catch (_) {}
}

// Hafızadaki kart sırasını geri yükle (geçerliyse), yoksa 0'dan başla
function restoreSavedIndex() {
    const saved = parseInt(localStorage.getItem(STORAGE_INDEX_KEY), 10);
    if (!isNaN(saved) && saved >= 0 && saved < cards.length) {
        currentIndex = saved;
    } else {
        currentIndex = 0;
    }
}

// Yüklenen kartları localStorage'a kaydet
function saveCardsToStorage(fileName) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(cards));
        if (fileName) localStorage.setItem(STORAGE_NAME_KEY, fileName);
    } catch (err) {
        console.warn("Kartlar hafızaya kaydedilemedi:", err);
    }
}

async function loadCards() {
    // Önce daha önce yüklenmiş bir dosya var mı diye bak
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            if (Array.isArray(parsed) && parsed.length > 0) {
                cards = parsed;
                restoreSavedIndex();
                loadVoices();
                const fileStatus = document.getElementById("file-status");
                const savedName = localStorage.getItem(STORAGE_NAME_KEY);
                if (fileStatus) {
                    fileStatus.textContent = savedName
                        ? `✅ ${cards.length} kelime (kayıtlı): ${savedName}`
                        : `✅ ${cards.length} kayıtlı kelime yüklendi`;
                }
                return;
            }
        } catch (err) {
            console.warn("Kayıtlı kartlar okunamadı, varsayılana dönülüyor:", err);
        }
    }
    // Kayıt yoksa varsayılan data.json'u yükle
    const response = await fetch("data.json");
    cards = await response.json();
    restoreSavedIndex();
    loadVoices();
}

function deduplicateCards(arr) {
    const seen = new Set();
    return arr.filter(card => {
        const key = (card.sentence || card.english || '').toLowerCase().trim();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

// ——— Excel / CSV yükleme ———
document.getElementById("file-input").addEventListener("change", function (e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    const fileStatus = document.getElementById("file-status");

    reader.onload = function (loadEvent) {
        try {
            let data;
            if (file.name.endsWith(".csv")) {
                data = XLSX.read(loadEvent.target.result, { type: "string" });
            } else {
                data = XLSX.read(loadEvent.target.result, { type: "array" });
            }

            const sheet = data.Sheets[data.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            const rows = json.filter(row => row.length >= 2 && row[0] && row[1]);

            if (rows.length === 0) {
                fileStatus.textContent = "❌ Dosyada geçerli veri bulunamadı!";
                return;
            }

            let parsed = rows.map(row => ({
                word: String(row[0]).trim().split(' ')[0].toLowerCase(),
                meanings: [String(row[1]).trim()],
                sentence: String(row[0]).trim(),
                turkish: String(row[1]).trim()
            }));

            const uniqueCount = parsed.length;
            parsed = deduplicateCards(parsed);
            cards = parsed;
            currentIndex = 0;
            isShowingExample = false;
            saveCardsToStorage(file.name);  // kalıcı olarak hafızaya kaydet
            showEnglish();

            const dedupMsg = uniqueCount !== cards.length
                ? ` (${uniqueCount - cards.length} tekrar temizlendi)`
                : "";
            fileStatus.textContent = `✅ ${cards.length} benzersiz kelime yüklendi${dedupMsg}: ${file.name}`;
        } catch (err) {
            fileStatus.textContent = "❌ Dosya okunamadı: " + err.message;
        }
    };

    if (file.name.endsWith(".csv")) {
        reader.readAsText(file, "UTF-8");
    } else {
        reader.readAsArrayBuffer(file);
    }
});

// ——— Sıfırla: yüklenen dosyayı sil, varsayılan listeye dön ———
const resetBtn = document.getElementById("reset-btn");
if (resetBtn) {
    resetBtn.addEventListener("click", async function () {
        if (!confirm("Yüklediğiniz dosya silinip varsayılan kelime listesine dönülecek. Emin misiniz?")) return;
        localStorage.removeItem(STORAGE_KEY);
        localStorage.removeItem(STORAGE_NAME_KEY);
        const response = await fetch("data.json");
        cards = await response.json();
        currentIndex = 0;
        isShowingExample = false;
        showEnglish();
        const fileStatus = document.getElementById("file-status");
        if (fileStatus) fileStatus.textContent = "Henüz dosya yüklenmedi";
    });
}

// =========================================================
// SES
// =========================================================
function loadVoices() {
    const voices = window.speechSynthesis.getVoices();
    if (voices.length > 0) {
        assignVoices(voices);
        showEnglish();
    } else {
        window.speechSynthesis.onvoiceschanged = () => {
            const allVoices = window.speechSynthesis.getVoices();
            assignVoices(allVoices);
            showEnglish();
        };
    }
}

function assignVoices(voices) {
    const englishVoices = voices.filter(v => v.lang.startsWith("en"));
    femaleVoice = englishVoices.find(v =>
        v.name.toLowerCase().includes("zira") ||
        v.name.toLowerCase().includes("female")
    );
    maleVoice = englishVoices.find(v =>
        v.name.toLowerCase().includes("david") ||
        v.name.toLowerCase().includes("mark") ||
        v.name.toLowerCase().includes("male")
    );
    if (!femaleVoice && englishVoices.length > 0) femaleVoice = englishVoices[0];
    if (!maleVoice && englishVoices.length > 1) maleVoice = englishVoices[1];
    if (!maleVoice && englishVoices.length > 0) maleVoice = englishVoices[0];
}

// Otomatik seslendirme (autoSpeak kapalıyken sessizdir)
function speakEnglish(text) {
    if (!autoSpeak) return; // otomatik seslendirme kapalıysa hiçbir şey yapma
    window.speechSynthesis.cancel();
    const speech = new SpeechSynthesisUtterance(text);
    speech.lang = "en-US";
    speech.rate = 0.85;
    speech.pitch = 1;
    const voice = currentIndex % 2 === 0 ? femaleVoice : maleVoice;
    if (voice) {
        speech.voice = voice;
    }
    window.speechSynthesis.speak(speech);
}

// ——— Geçerli kartın İngilizce cümlesini döndür (örnek modunda örnek cümle) ———
function getCurrentSentenceText() {
    if (isShowingExample) {
        const h2 = document.querySelector("#card h2");
        return h2 ? h2.textContent : "";
    }
    if (cards[currentIndex]) {
        return cards[currentIndex].sentence || cards[currentIndex].english || "";
    }
    return "";
}

// ——— İsteğe bağlı (elle) seslendirme — 🔊 butonu bunu çağırır ———
function speakCurrentSentence() {
    const text = getCurrentSentenceText();
    if (!text) return;

    window.speechSynthesis.cancel();
    const speech = new SpeechSynthesisUtterance(text);
    speech.lang = "en-US";
    speech.rate = 0.85;
    speech.pitch = 1;
    const voice = currentIndex % 2 === 0 ? femaleVoice : maleVoice;
    if (voice) speech.voice = voice;
    window.speechSynthesis.speak(speech);
}

// =========================================================
// SAYAÇ
// =========================================================
function updateCounter() {
    const counter = document.getElementById("counter");
    if (cards.length > 0) {
        if (isShowingExample) {
            counter.textContent = `📝 Örnek · ${originalCardIndex + 1} / ${cards.length}`;
        } else {
            const card = cards[currentIndex];
            const word = card.word || '';
            counter.textContent = `${word} · ${currentIndex + 1} / ${cards.length}`;
        }
    } else {
        counter.textContent = "";
    }
}

// =========================================================
// ZAMANLAYICI YARDIMCILARI
// =========================================================
function clearTimers() {
    clearTimeout(hideTooltipTimer);
    clearTimeout(hoverTimer);
}

function scheduleHideTooltip(delay) {
    clearTimeout(hideTooltipTimer);
    hideTooltipTimer = setTimeout(() => {
        const tooltip = document.getElementById("word-tooltip");
        if (!tooltip || !tooltip.matches(':hover')) {
            hideTooltip();
        }
    }, delay);
}

// =========================================================
// İNGİLİZCE KALIPLAR (PHRASES) — cümlede tek tıklanabilir birim olarak
// =========================================================
const PHRASES = [
    // Bağlaç / edat kalıpları
    "as long as", "as soon as", "as well as", "as far as", "as much as",
    "as if", "as though", "as a result of", "as a result", "as opposed to",
    "in case of", "in case", "in order to", "in order for", "in spite of",
    "in terms of", "in addition to", "in addition", "in front of",
    "in favor of", "in favour of", "in charge of", "in accordance with",
    "in advance", "in line with", "in light of", "in the long run",
    "in the short term", "in the event of", "in response to", "in contrast to",
    "instead of", "the case for", "the case against",
    "due to", "owing to", "thanks to", "according to", "regardless of",
    "apart from", "aside from", "rather than", "other than", "such as",
    "so that", "so as to", "even though", "even if", "for the sake of",
    "on behalf of", "on account of", "with respect to", "with regard to",
    "with regards to", "by means of", "ahead of", "prior to", "subject to",
    "based on", "up to date", "more or less", "at least", "at most",
    "at first", "at last", "no longer", "as of",
    // Fiil kalıpları (phrasal / collocation)
    "look forward to", "take into account", "take advantage of",
    "give rise to", "come up with", "carry out", "point out", "set up",
    "follow up", "break down", "deal with", "rely on", "depend on",
    "result in", "lead to", "focus on", "go through", "bring about",
    "phase out", "roll out", "scale up", "ramp up", "follow through",
    "keep up with", "catch up with", "make up for", "account for",
    "due diligence", "supply chain", "value chain", "market share",
    "go to market", "first mover", "first-mover advantage", "first mover advantage",
    // ——— Öbek fiiller (phrasal verbs) — birleşik halde anlam ifade edenler ———
    // Üç kelimeli olanlar (en uzun önce eşleşsin diye listede önde)
    "look forward to", "get along with", "come up against", "put up with",
    "look down on", "look up to", "get away with", "run out of",
    "cut down on", "drop out of", "stand up for", "face up to",
    "live up to", "get on with", "go back on", "fall back on",
    "keep away from", "look out for", "watch out for", "come down with",
    // İki kelimeli öbek fiiller
    "roll up", "roll back", "roll over", "back up", "back down",
    "break up", "break in", "break out", "break off", "break through",
    "bring up", "bring in", "bring out", "bring down", "bring back",
    "call off", "call up", "call back", "carry on", "check in",
    "check out", "come across", "come back", "come in", "come out",
    "come over", "come along", "cut off", "cut out", "cut back",
    "drop off", "drop out", "end up", "fall apart", "fall behind",
    "fill in", "fill out", "fill up", "find out", "get up",
    "get back", "get in", "get off", "get on", "get out",
    "get over", "get through", "give up", "give in", "give away",
    "give back", "go ahead", "go on", "go off", "go out",
    "go over", "grow up", "hand in", "hand out", "hand over",
    "hang on", "hang up", "hold on", "hold back", "hold up",
    "keep on", "keep up", "lay off", "let down", "log in",
    "log out", "look after", "look into", "look up", "make out",
    "make up", "move on", "pass away", "pass out", "pay off",
    "pay back", "pick up", "pick out", "point out", "pull off",
    "pull out", "pull over", "put off", "put on", "put out",
    "put up", "run into", "run out", "run over", "set off",
    "set out", "set aside", "settle down", "show up", "shut down",
    "sign in", "sign up", "sign out", "slow down", "sort out",
    "speed up", "stand out", "stand by", "stay up", "step down",
    "switch on", "switch off", "take off", "take on", "take over",
    "take up", "take back", "tear down", "throw away", "throw out",
    "try out", "turn on", "turn off", "turn up", "turn down",
    "turn out", "turn over", "turn into", "wake up", "warm up",
    "wear out", "work out", "wrap up", "write down", "write up",
    "back off", "blow up", "calm down", "carry off", "catch up",
    "clear up", "close down", "count on", "draw up", "dry up",
    "figure out", "hand back", "kick off", "knock out", "lay out",
    "leave out", "let in", "line up", "look back", "map out",
    "open up", "pile up", "play down", "point to", "push back",
    "rule out", "scale back", "screen out", "set back", "shake up",
    "step up", "stock up", "sum up", "team up", "tie up",
    "tone down", "track down", "use up", "weigh in", "win over",
    "zoom in", "zoom out"
];

// =========================================================
// KALIPLARIN GÖMÜLÜ TÜRKÇE SÖZLÜĞÜ
// İnternet/çeviri servisi olmadan da kalıp anlamları anında gösterilir.
// Birden çok anlam virgül/noktalı virgülle ayrılır.
// =========================================================
const PHRASE_TR = {
    // Bağlaç / edat kalıpları
    "as long as": "-dığı sürece, yeter ki",
    "as soon as": "-er -mez, olur olmaz",
    "as well as": "-nın yanı sıra, ayrıca, -e ek olarak",
    "as far as": "-e kadar; -e gelince, bildiği kadarıyla",
    "as much as": "-dığı kadar, o kadar",
    "as if": "sanki, -mış gibi",
    "as though": "sanki, -mış gibi",
    "as a result of": "-in sonucunda, -den dolayı",
    "as a result": "sonuç olarak, bunun sonucunda",
    "as opposed to": "-in aksine, -e karşılık",
    "in case of": "durumunda, halinde",
    "in case": "ihtimaline karşı, belki diye",
    "in order to": "-mek için, amacıyla",
    "in order for": "-in ... yapabilmesi için",
    "in spite of": "-e rağmen, -e karşın",
    "in terms of": "açısından, bakımından",
    "in addition to": "-e ek olarak, -in yanı sıra",
    "in addition": "ayrıca, ek olarak, üstelik",
    "in front of": "-in önünde",
    "in favor of": "-in lehine, -den yana",
    "in favour of": "-in lehine, -den yana",
    "in charge of": "-den sorumlu, -in başında",
    "in accordance with": "-e uygun olarak, uyarınca",
    "in advance": "önceden, peşinen",
    "in line with": "-e paralel olarak, -le uyumlu",
    "in light of": "ışığında, göz önüne alındığında",
    "in the long run": "uzun vadede, eninde sonunda",
    "in the short term": "kısa vadede",
    "in the event of": "olması halinde, durumunda",
    "in response to": "-e yanıt olarak, karşılık olarak",
    "in contrast to": "-in aksine, -e kıyasla",
    "instead of": "yerine, -ecek yerde",
    "the case for": "lehine gerekçeler, -i savunan görüş",
    "the case against": "aleyhine gerekçeler, -e karşı görüş",
    "due to": "-den dolayı, nedeniyle, yüzünden",
    "owing to": "nedeniyle, -den ötürü",
    "thanks to": "sayesinde",
    "according to": "-e göre",
    "regardless of": "-e bakılmaksızın, ne olursa olsun",
    "apart from": "-den başka, dışında; bir yana",
    "aside from": "-in dışında, -den başka",
    "rather than": "yerine, -mektense",
    "other than": "-den başka, dışında",
    "such as": "gibi, örneğin",
    "so that": "-sın diye, böylece",
    "so as to": "-mek için, amacıyla",
    "even though": "-e rağmen, her ne kadar",
    "even if": "-se bile",
    "for the sake of": "hatırına, uğruna, iyiliği için",
    "on behalf of": "adına, namına",
    "on account of": "yüzünden, nedeniyle",
    "with respect to": "-e ilişkin, açısından",
    "with regard to": "-e gelince, ile ilgili olarak",
    "with regards to": "-e gelince, ile ilgili olarak",
    "by means of": "aracılığıyla, vasıtasıyla",
    "ahead of": "-den önce, -in önünde",
    "prior to": "-den önce, öncesinde",
    "subject to": "-e tabi, -e bağlı, koşuluyla",
    "based on": "-e dayalı, temel alınarak, -e dayanarak",
    "up to date": "güncel, çağdaş",
    "more or less": "aşağı yukarı, az çok",
    "at least": "en azından, hiç değilse",
    "at most": "en fazla, olsa olsa",
    "at first": "ilk başta, önceleri",
    "at last": "sonunda, nihayet",
    "no longer": "artık ... değil, bundan böyle ... -me-",
    "as of": "itibarıyla, -den itibaren",
    // Fiil kalıpları (phrasal / collocation)
    "look forward to": "dört gözle beklemek, sabırsızlıkla beklemek",
    "take into account": "hesaba katmak, göz önünde bulundurmak",
    "take advantage of": "-den yararlanmak; istismar etmek",
    "give rise to": "yol açmak, neden olmak, doğurmak",
    "come up with": "(fikir/çözüm) bulmak, ortaya atmak",
    "carry out": "yürütmek, gerçekleştirmek, yerine getirmek",
    "point out": "belirtmek, işaret etmek, dikkat çekmek",
    "set up": "kurmak, düzenlemek, ayarlamak",
    "follow up": "takip etmek, izlemek, devamını getirmek",
    "break down": "bozulmak; ayrıştırmak; sinir krizi geçirmek",
    "deal with": "ilgilenmek, başa çıkmak, ele almak",
    "rely on": "güvenmek, bel bağlamak",
    "depend on": "-e bağlı olmak; güvenmek",
    "result in": "-le sonuçlanmak, yol açmak",
    "lead to": "-e yol açmak, neden olmak",
    "focus on": "odaklanmak, yoğunlaşmak",
    "go through": "yaşamak, geçirmek; gözden geçirmek",
    "bring about": "meydana getirmek, yol açmak",
    "phase out": "aşamalı olarak kaldırmak, kademeli sonlandırmak",
    "roll out": "piyasaya sürmek, yaygınlaştırmak",
    "scale up": "ölçeği büyütmek, artırmak",
    "ramp up": "hızla artırmak, yükseltmek",
    "follow through": "sonuna kadar götürmek, tamamlamak",
    "keep up with": "ayak uydurmak, geride kalmamak",
    "catch up with": "yetişmek, arayı kapatmak",
    "make up for": "telafi etmek, karşılamak",
    "account for": "açıklamak; (pay) oluşturmak",
    "due diligence": "durum tespiti, özenli inceleme",
    "supply chain": "tedarik zinciri",
    "value chain": "değer zinciri",
    "market share": "pazar payı",
    "go to market": "pazara açılma, pazara giriş",
    "first mover": "ilk hamleyi yapan, öncü",
    "first-mover advantage": "ilk hamle avantajı, öncü olma avantajı",
    "first mover advantage": "ilk hamle avantajı, öncü olma avantajı",
    // Üç kelimeli öbek fiiller
    "get along with": "iyi geçinmek, anlaşmak",
    "come up against": "(zorlukla) karşılaşmak",
    "put up with": "katlanmak, tahammül etmek",
    "look down on": "küçümsemek, hor görmek",
    "look up to": "saygı duymak, örnek almak",
    "get away with": "yanına kâr kalmak, paçayı kurtarmak",
    "run out of": "-i bitirmek, -i tüketmek",
    "cut down on": "azaltmak, kısmak",
    "drop out of": "yarıda bırakmak, (okulu) terk etmek",
    "stand up for": "savunmak, arkasında durmak",
    "face up to": "yüzleşmek, göğüs germek",
    "live up to": "(beklentiyi) karşılamak, layık olmak",
    "get on with": "devam etmek; iyi geçinmek",
    "go back on": "sözünden dönmek, caymak",
    "fall back on": "-e başvurmak, yedeğine sarılmak",
    "keep away from": "uzak durmak, yaklaşmamak",
    "look out for": "kollamak, gözetmek; dikkat etmek",
    "watch out for": "dikkat etmek, gözünü açmak",
    "come down with": "(hastalığa) yakalanmak",
    // İki kelimeli öbek fiiller
    "roll up": "sarmak, kıvırmak; toplamak",
    "roll back": "geri çekmek, (fiyat) düşürmek",
    "roll over": "yuvarlanmak; devretmek",
    "back up": "yedeklemek; desteklemek",
    "back down": "geri adım atmak, vazgeçmek",
    "break up": "ayrılmak; dağıtmak, dağılmak",
    "break in": "zorla girmek; alıştırmak",
    "break out": "patlak vermek; kaçmak",
    "break off": "koparmak; aniden kesmek",
    "break through": "yarıp geçmek; çığır açmak",
    "bring up": "gündeme getirmek; (çocuk) yetiştirmek",
    "bring in": "getirmek, kazandırmak",
    "bring out": "ortaya çıkarmak; piyasaya çıkarmak",
    "bring down": "düşürmek, indirmek, devirmek",
    "bring back": "geri getirmek; hatırlatmak",
    "call off": "iptal etmek, vazgeçmek",
    "call up": "telefon etmek; askere çağırmak",
    "call back": "geri aramak",
    "carry on": "devam etmek, sürdürmek",
    "check in": "giriş yapmak, kayıt yaptırmak",
    "check out": "çıkış yapmak; incelemek, göz atmak",
    "come across": "rastlamak, denk gelmek; izlenim bırakmak",
    "come back": "geri dönmek",
    "come in": "girmek, içeri gelmek",
    "come out": "ortaya çıkmak; yayımlanmak",
    "come over": "uğramak, ziyarete gelmek",
    "come along": "birlikte gelmek; ilerlemek",
    "cut off": "kesmek, ayırmak; sözünü kesmek",
    "cut out": "kesip çıkarmak; bırakmak",
    "cut back": "kısmak, azaltmak",
    "drop off": "(birini) bırakmak; azalmak; uyuyakalmak",
    "drop out": "(okulu) bırakmak, çekilmek",
    "end up": "sonunda ... olmak, -le sonuçlanmak",
    "fall apart": "dağılmak, parçalanmak",
    "fall behind": "geri kalmak, geride kalmak",
    "fill in": "(form) doldurmak; yerine bakmak",
    "fill out": "(form) doldurmak",
    "fill up": "doldurmak, dolmak",
    "find out": "öğrenmek, ortaya çıkarmak",
    "get up": "kalkmak, yataktan kalkmak",
    "get back": "geri dönmek; geri almak",
    "get in": "girmek, (araca) binmek",
    "get off": "inmek; kurtulmak",
    "get on": "binmek; ilerlemek, iyi gitmek",
    "get out": "çıkmak, dışarı çıkmak",
    "get over": "atlatmak, üstesinden gelmek",
    "get through": "tamamlamak, atlatmak; (telefonda) ulaşmak",
    "give up": "vazgeçmek, bırakmak, pes etmek",
    "give in": "boyun eğmek, pes etmek",
    "give away": "bağışlamak; ele vermek",
    "give back": "geri vermek, iade etmek",
    "go ahead": "devam etmek; önden gitmek; buyurun",
    "go on": "devam etmek; olmak, yaşanmak",
    "go off": "patlamak; (alarm) çalmak; bozulmak",
    "go out": "dışarı çıkmak; sönmek",
    "go over": "gözden geçirmek, üzerinden geçmek",
    "grow up": "büyümek, yetişkin olmak",
    "hand in": "teslim etmek, vermek",
    "hand out": "dağıtmak",
    "hand over": "devretmek, teslim etmek",
    "hang on": "beklemek; sıkı tutunmak",
    "hang up": "telefonu kapatmak; asmak",
    "hold on": "beklemek; tutunmak",
    "hold back": "geride tutmak, kendini tutmak",
    "hold up": "geciktirmek; dayanmak; soymak",
    "keep on": "devam etmek, sürdürmek",
    "keep up": "sürdürmek; ayak uydurmak",
    "lay off": "işten çıkarmak",
    "let down": "hayal kırıklığına uğratmak, yüzüstü bırakmak",
    "log in": "oturum açmak, giriş yapmak",
    "log out": "oturumu kapatmak, çıkış yapmak",
    "look after": "bakmak, ilgilenmek",
    "look into": "araştırmak, incelemek",
    "look up": "(sözlükten) bakmak; iyiye gitmek",
    "make out": "seçebilmek, güçlükle anlamak",
    "make up": "uydurmak; barışmak; oluşturmak",
    "move on": "yoluna devam etmek, (yeni konuya) geçmek",
    "pass away": "vefat etmek, hayatını kaybetmek",
    "pass out": "bayılmak",
    "pay off": "karşılığını vermek; (borcu) kapatmak",
    "pay back": "geri ödemek",
    "pick up": "almak, yerden kaldırmak; (birini) almak; öğrenivermek",
    "pick out": "seçmek, ayırt etmek",
    "pull off": "(zor işi) başarmak",
    "pull out": "çekilmek, vazgeçmek",
    "pull over": "(aracı) kenara çekmek",
    "put off": "ertelemek",
    "put on": "giymek; (kilo) almak",
    "put out": "söndürmek",
    "put up": "asmak; barındırmak; (fiyat) artırmak",
    "run into": "rastlamak; çarpmak",
    "run out": "tükenmek, bitmek",
    "run over": "(araçla) ezmek; üzerinden geçmek",
    "set off": "yola çıkmak; tetiklemek",
    "set out": "yola koyulmak; ortaya koymak",
    "set aside": "ayırmak, bir kenara koymak",
    "settle down": "yerleşmek, durulmak",
    "show up": "ortaya çıkmak, gelmek",
    "shut down": "kapatmak, kapanmak",
    "sign in": "giriş yapmak, oturum açmak",
    "sign up": "kaydolmak, üye olmak",
    "sign out": "çıkış yapmak, oturumu kapatmak",
    "slow down": "yavaşlamak, yavaşlatmak",
    "sort out": "halletmek, çözmek; ayıklamak",
    "speed up": "hızlandırmak, hızlanmak",
    "stand out": "öne çıkmak, göze çarpmak",
    "stand by": "hazır beklemek; arkasında durmak",
    "stay up": "geç saate kadar uyanık kalmak",
    "step down": "istifa etmek, görevi bırakmak",
    "switch on": "(cihazı) açmak",
    "switch off": "(cihazı) kapatmak",
    "take off": "havalanmak; (giysi) çıkarmak; hızla yükselmek",
    "take on": "üstlenmek; işe almak",
    "take over": "devralmak, ele geçirmek",
    "take up": "(hobiye) başlamak; yer kaplamak",
    "take back": "geri almak; sözünü geri almak",
    "tear down": "yıkmak",
    "throw away": "çöpe atmak, elden çıkarmak",
    "throw out": "atmak; kovmak",
    "try out": "denemek, sınamak",
    "turn on": "açmak (ışık/cihaz)",
    "turn off": "kapatmak (ışık/cihaz)",
    "turn up": "ortaya çıkmak, çıkagelmek; (sesi) açmak",
    "turn down": "reddetmek; (sesi) kısmak",
    "turn out": "ortaya çıkmak, ... olduğu anlaşılmak",
    "turn over": "ters çevirmek; devretmek; ciro yapmak",
    "turn into": "-e dönüşmek, -e dönüştürmek",
    "wake up": "uyanmak, uyandırmak",
    "warm up": "ısınmak, ısıtmak",
    "wear out": "eskitmek, eskimek; yormak",
    "work out": "çözmek; spor yapmak; yolunda gitmek",
    "wrap up": "tamamlamak, bitirmek; sarmak",
    "write down": "not etmek, yazmak",
    "write up": "kaleme almak, rapor haline getirmek",
    "back off": "geri çekilmek, rahat bırakmak",
    "blow up": "patlamak, patlatmak; şişirmek",
    "calm down": "sakinleşmek, sakinleştirmek",
    "carry off": "başarıyla üstesinden gelmek; alıp götürmek",
    "catch up": "yetişmek, arayı kapatmak",
    "clear up": "açıklığa kavuşturmak; (hava) açmak",
    "close down": "(işyerini) kapatmak, kapanmak",
    "count on": "güvenmek, bel bağlamak",
    "draw up": "(belge) hazırlamak, kaleme almak",
    "dry up": "kurumak, kuruyup tükenmek",
    "figure out": "çözmek, anlamak, akıl erdirmek",
    "hand back": "geri vermek",
    "kick off": "başlamak, başlatmak",
    "knock out": "nakavt etmek, bayıltmak; elemek",
    "lay out": "sermek; düzenlemek, tasarlamak",
    "leave out": "dışarıda bırakmak, atlamak",
    "let in": "içeri almak",
    "line up": "sıraya girmek, sıralamak; ayarlamak",
    "look back": "geriye dönüp bakmak",
    "map out": "ayrıntılı planlamak",
    "open up": "açılmak; içini dökmek",
    "pile up": "birikmek, yığılmak",
    "play down": "önemsiz göstermek, hafife almak",
    "point to": "işaret etmek, göstermek",
    "push back": "geri itmek; karşı çıkmak; ertelemek",
    "rule out": "olasılık dışı bırakmak, elemek",
    "scale back": "küçültmek, azaltmak",
    "screen out": "elemek, ayıklamak",
    "set back": "geciktirmek, sekteye uğratmak",
    "shake up": "sarsmak; köklü değişiklik yapmak",
    "step up": "artırmak, hızlandırmak; öne çıkmak",
    "stock up": "stok yapmak, depolamak",
    "sum up": "özetlemek, toparlamak",
    "team up": "güçlerini birleştirmek, takım kurmak",
    "tie up": "bağlamak; meşgul etmek",
    "tone down": "yumuşatmak, dozunu azaltmak",
    "track down": "izini sürüp bulmak",
    "use up": "tüketmek, bitirmek",
    "weigh in": "görüş bildirmek, tartışmaya katılmak",
    "win over": "(birini) kazanmak, ikna etmek",
    "zoom in": "yakınlaştırmak",
    "zoom out": "uzaklaştırmak"
};

let _phraseRegex = null;
function getPhraseRegex() {
    if (_phraseRegex) return _phraseRegex;
    const sorted = [...PHRASES].sort((a, b) => b.length - a.length);
    const escaped = sorted.map(p => p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    _phraseRegex = new RegExp('\\b(?:' + escaped.join('|') + ')\\b', 'gi');
    return _phraseRegex;
}

// Cümleyi kalıpları tek parça, geri kalanı verilen tokenizer ile işleyerek HTML üretir
function buildHtmlWithPhrases(sentence, tokenizeSegment) {
    const regex = getPhraseRegex();
    regex.lastIndex = 0;
    let result = '';
    let lastIndex = 0;
    let m;
    while ((m = regex.exec(sentence)) !== null) {
        result += tokenizeSegment(sentence.slice(lastIndex, m.index));
        const phraseText = m[0];
        result += `<span class="word-phrase" data-phrase="${phraseText}" data-lower="${phraseText.toLowerCase()}">${phraseText}</span>`;
        lastIndex = m.index + phraseText.length;
        if (phraseText.length === 0) regex.lastIndex++; // güvenlik
    }
    result += tokenizeSegment(sentence.slice(lastIndex));
    return result;
}

// Düz metin parçasını kelimelere ayırıp anahtar kelimeyi vurgular
function tokenizeSegmentKeyword(text, keyWord) {
    const tokens = text.match(/[\w']+(?:-[\w']+)*|[^\w']+/g) || [];
    return tokens.map(token => {
        if (keyWord && token.toLowerCase() === keyWord.toLowerCase()) {
            return `<span class="word-keyword" data-word="${token}" data-lower="${token.toLowerCase()}">${token}</span>`;
        }
        if (/^[a-zA-Z][a-zA-Z'-]*$/.test(token)) {
            return `<span class="word-clickable" data-word="${token}" data-lower="${token.toLowerCase()}">${token}</span>`;
        }
        return token;
    }).join('');
}

// Düz metin parçasını kelimelere ayırıp eşleşen kelimeyi (örnek kartında) vurgular
function tokenizeSegmentHighlight(text, word) {
    const tokens = text.match(/[\w']+(?:-[\w']+)*|[^\w']+/g) || [];
    return tokens.map(token => {
        if (word && token.toLowerCase() === word.toLowerCase()) {
            return `<span class="word-highlight">${token}</span>`;
        }
        if (/^[a-zA-Z][a-zA-Z'-]*$/.test(token)) {
            return `<span class="word-clickable" data-word="${token}" data-lower="${token.toLowerCase()}">${token}</span>`;
        }
        return token;
    }).join('');
}

// =========================================================
// ANA EKRAN – İNGİLİZCE CÜMLE + ANAHTAR KELİME VURGULU
// =========================================================
function showEnglish() {
    showingTurkish = false;
    const card = document.getElementById("card");
    hideTooltip();
    clearTimers();
    selAnchorIdx = null; // kart yeniden çiziliyor → yarım kalan seçim iptal

    if (isShowingExample) {
        showOriginalCard();
        return;
    }

    const cardData = cards[currentIndex];
    const sentence = cardData.sentence || cardData.english || '';
    const keyWord = cardData.word || '';

    // Cümleyi kalıplar (örn. "as long as") tek parça, geri kalanı kelime kelime işle
    const html = buildHtmlWithPhrases(sentence, (t) => tokenizeSegmentKeyword(t, keyWord));

    card.innerHTML = `<h2>${html}</h2>`;

    // Tıklanabilir kelimelere olay dinleyicisi ekle
    attachWordListeners(card);

    speakEnglish(sentence);
    updateCounter();
    saveCurrentIndex();   // kaldığımız kartı hafızaya yaz
}

function attachWordListeners(card) {
    // Anahtar kelime (vurgulu) - özel davranış: Türkçe anlamları göster
    card.querySelectorAll('.word-keyword').forEach(span => {
        span.addEventListener('click', function (e) {
            e.stopPropagation();
            clearTimers();
            if (selectMode) { handleSelectTap(this); return; }
            // Aynı kelimeye tekrar basılırsa açıklamayı kapat (aç/kapa)
            if (activeTooltipEl === this && document.getElementById('word-tooltip')) {
                hideTooltip();
                return;
            }
            const word = this.dataset.word;
            // Anahtar kelimeye tıklandı → karttaki Türkçe anlamları göster
            showKeywordMeanings(word, this);
        });

        span.addEventListener('mouseenter', function () {
            if (selectMode) return;
            clearTimeout(hoverTimer);
            hoverTimer = setTimeout(() => {
                const word = this.dataset.word;
                showKeywordMeanings(word, this);
            }, 300);
        });

        span.addEventListener('mouseleave', function () {
            clearTimeout(hoverTimer);
            if (selectMode) return;
            scheduleHideTooltip(400);
        });
    });

    // Diğer tıklanabilir kelimeler - dictionary API ile çeviri
    card.querySelectorAll('.word-clickable').forEach(span => {
        span.addEventListener('click', function (e) {
            e.stopPropagation();
            clearTimers();
            if (selectMode) { handleSelectTap(this); return; }
            // Aynı kelimeye tekrar basılırsa açıklamayı kapat (aç/kapa)
            if (activeTooltipEl === this && document.getElementById('word-tooltip')) {
                hideTooltip();
                return;
            }
            const word = this.dataset.word;
            const lower = this.dataset.lower;
            showWordMeanings(lower, word, this);
        });

        span.addEventListener('mouseenter', function () {
            if (selectMode) return;
            clearTimeout(hoverTimer);
            hoverTimer = setTimeout(() => {
                const word = this.dataset.word;
                const lower = this.dataset.lower;
                showWordMeanings(lower, word, this);
            }, 300);
        });

        span.addEventListener('mouseleave', function () {
            clearTimeout(hoverTimer);
            if (selectMode) return;
            scheduleHideTooltip(400);
        });
    });

    // Kalıplar (phrase) - kalıbın Türkçe anlamını göster
    card.querySelectorAll('.word-phrase').forEach(span => {
        span.addEventListener('click', function (e) {
            e.stopPropagation();
            clearTimers();
            if (selectMode) { handleSelectTap(this); return; }
            // Aynı kalıba tekrar basılırsa açıklamayı kapat (aç/kapa)
            if (activeTooltipEl === this && document.getElementById('word-tooltip')) {
                hideTooltip();
                return;
            }
            showPhraseMeaning(this.dataset.phrase, this);
        });

        span.addEventListener('mouseenter', function () {
            if (selectMode) return;
            clearTimeout(hoverTimer);
            hoverTimer = setTimeout(() => {
                showPhraseMeaning(this.dataset.phrase, this);
            }, 300);
        });

        span.addEventListener('mouseleave', function () {
            clearTimeout(hoverTimer);
            if (selectMode) return;
            scheduleHideTooltip(400);
        });
    });
}

// =========================================================
// SEÇİM MODU — cümlede kelime / kelime grubu seçme
// İlk dokunuş başlangıcı, ikinci dokunuş bitişi işaretler;
// aradaki tüm sözcükler (noktalama dahil) tek parça çevrilir.
// =========================================================
function getCardTokens() {
    const card = document.getElementById("card");
    return Array.from(card.querySelectorAll('.word-clickable, .word-keyword, .word-phrase'));
}

function clearSelection() {
    selAnchorIdx = null;
    getCardTokens().forEach(t => t.classList.remove('word-selected', 'word-sel-anchor'));
}

function handleSelectTap(el) {
    const tokens = getCardTokens();
    const idx = tokens.indexOf(el);
    if (idx < 0) return;

    // İlk dokunuş → başlangıç noktasını işaretle
    if (selAnchorIdx === null) {
        clearSelection();
        hideTooltip();
        selAnchorIdx = idx;
        el.classList.add('word-selected', 'word-sel-anchor');
        showToast("Şimdi son kelimeye dokun · tek kelime için aynısına tekrar dokun");
        return;
    }

    // İkinci dokunuş → aralığı vurgula ve çevir
    const start = Math.min(selAnchorIdx, idx);
    const end = Math.max(selAnchorIdx, idx);
    selAnchorIdx = null;

    tokens.forEach((t, i) => {
        t.classList.toggle('word-selected', i >= start && i <= end);
        t.classList.remove('word-sel-anchor');
    });

    // Tek kelime seçildi → normal zengin balonu göster
    if (start === end) {
        const t = tokens[start];
        if (t.classList.contains('word-phrase')) {
            showPhraseMeaning(t.dataset.phrase, t);
        } else if (t.classList.contains('word-keyword')) {
            showKeywordMeanings(t.dataset.word, t);
        } else {
            showWordMeanings(t.dataset.lower, t.dataset.word, t);
        }
        return;
    }

    // Ekranda görünen metni aralık olarak al (aradaki noktalama/boşluklar dahil)
    const range = document.createRange();
    range.setStartBefore(tokens[start]);
    range.setEndAfter(tokens[end]);
    const text = range.toString().replace(/\s+/g, " ").trim();
    showSelectionTooltip(text, tokens[end]);
}

function showSelectionTooltip(text, anchorEl) {
    hideTooltip();

    const tooltip = document.createElement("div");
    tooltip.className = "word-tooltip";
    tooltip.id = "word-tooltip";
    tooltip.innerHTML = `<div class="tooltip-header">
<strong>${text}</strong>
<span class="tooltip-tr">🖍 seçim</span>
</div>
<div class="tooltip-meanings">
<div class="meaning-group">
<div class="meaning-pos">🇹🇷 TÜRKÇE KARŞILIĞI</div>
<div class="meaning-item">
<div class="meaning-def" data-seltr>🔄 çevriliyor...</div>
<div class="meaning-def-tr" data-selalt hidden></div>
</div>
</div>
</div>`;
    document.body.appendChild(tooltip);
    positionTooltip(tooltip, anchorEl);

    tooltip.addEventListener('mouseenter', () => clearTimeout(hideTooltipTimer));
    tooltip.addEventListener('mouseleave', () => scheduleHideTooltip(300));
    activeTooltipEl = anchorEl;

    (async () => {
        const d = await fetchTurkishDict(text);
        if (!document.body.contains(tooltip)) return;
        const el = tooltip.querySelector('[data-seltr]');
        const altEl = tooltip.querySelector('[data-selalt]');
        if (el) {
            el.textContent = d.main
                ? `🇹🇷 ${d.main}`
                : "Çeviriye ulaşılamadı — tekrar deneyin";
        }
        const extras = dictExtraTerms(d);
        if (altEl && extras.length > 0) {
            altEl.hidden = false;
            altEl.textContent = `Diğer karşılıklar: ${extras.join(", ")}`;
        }
    })();
}

// ——— Kısa bilgi mesajı (toast) ———
let toastTimer = null;
function showToast(msg) {
    let el = document.getElementById("app-toast");
    if (!el) {
        el = document.createElement("div");
        el.id = "app-toast";
        el.className = "app-toast";
        document.body.appendChild(el);
    }
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove("show"), 2600);
}

// =========================================================
// TOOLTIP KONUMLANDIRMA (ortak yardımcı)
// =========================================================
function positionTooltip(tooltip, spanElement) {
    const rect = spanElement.getBoundingClientRect();
    tooltip.style.opacity = "0";
    const tooltipRect = tooltip.getBoundingClientRect();
    let left = rect.left + rect.width / 2 - tooltipRect.width / 2;
    let top = rect.top - tooltipRect.height - 12;
    if (left < 8) left = 8;
    if (left + tooltipRect.width > window.innerWidth - 8) {
        left = window.innerWidth - tooltipRect.width - 8;
    }
    if (top < 8) {
        top = rect.bottom + 12;
    }
    tooltip.style.left = left + "px";
    tooltip.style.top = top + "px";
    requestAnimationFrame(() => {
        tooltip.style.opacity = "1";
    });
}

// ——— Tıklanır tıklanmaz görünen "yükleniyor" balonu ———
function showLoadingTooltip(spanElement, word) {
    hideTooltip();

    const tooltip = document.createElement("div");
    tooltip.className = "word-tooltip";
    tooltip.id = "word-tooltip";
    tooltip.innerHTML = `<div class="tooltip-header">
<strong>${word}</strong>
<span class="tooltip-tr loading">🔄 çevriliyor...</span>
</div>`;
    document.body.appendChild(tooltip);
    positionTooltip(tooltip, spanElement);

    tooltip.addEventListener('mouseenter', () => clearTimeout(hideTooltipTimer));
    tooltip.addEventListener('mouseleave', () => scheduleHideTooltip(300));

    activeTooltipEl = spanElement;
}

// =========================================================
// KALIBIN (PHRASE) TÜRKÇE ANLAMINI GÖSTER
// =========================================================
function showPhraseMeaning(phrase, spanElement) {
    hideTooltip();

    const tooltip = document.createElement("div");
    tooltip.className = "word-tooltip";
    tooltip.id = "word-tooltip";

    tooltip.innerHTML = `<div class="tooltip-header">
<strong>${phrase}</strong>
<span class="tooltip-tr">🧩 kalıp</span>
</div>
<div class="tooltip-meanings">
<div class="meaning-group">
<div class="meaning-pos">🧩 KALIP ANLAMI</div>
<div class="meaning-item">
<div class="meaning-def" data-phrasetr>🔄 çevriliyor...</div>
<div class="meaning-def-tr" data-phrasealt hidden></div>
</div>
</div>
</div>`;

    document.body.appendChild(tooltip);
    positionTooltip(tooltip, spanElement);

    tooltip.addEventListener('mouseenter', () => clearTimeout(hideTooltipTimer));
    tooltip.addEventListener('mouseleave', () => scheduleHideTooltip(300));

    activeTooltipEl = spanElement;

    // Kalıbın Türkçe karşılıklarını getir (ana anlam + varsa diğer anlamlar)
    (async () => {
        const d = await fetchTurkishDict(phrase);
        if (!document.body.contains(tooltip)) return;
        const el = tooltip.querySelector('[data-phrasetr]');
        const altEl = tooltip.querySelector('[data-phrasealt]');
        if (el) {
            el.textContent = d.main
                ? `🇹🇷 ${d.main}`
                : 'Çeviriye ulaşılamadı — kalıba tekrar dokunarak yeniden deneyin';
        }
        const extras = dictExtraTerms(d);
        if (altEl && extras.length > 0) {
            altEl.hidden = false;
            altEl.textContent = `Diğer karşılıklar: ${extras.join(", ")}`;
        }
    })();
}

function showTurkish() {
    window.speechSynthesis.cancel();
    showingTurkish = true;
    const card = document.getElementById("card");
    hideTooltip();
    clearTimers();
    selAnchorIdx = null;

    const cardData = cards[currentIndex];
    const turkishText = cardData.turkish || '';

    card.innerHTML = `<p>${turkishText}</p>`;
}

// =========================================================
// ANAHTAR KELİMENİN TÜRKÇE ANLAMLARINI GÖSTER (data.json'dan)
// =========================================================
function showKeywordMeanings(word, spanElement) {
    hideTooltip();

    const cardData = cards[currentIndex];
    const meanings = cardData.meanings || [];

    const tooltip = document.createElement("div");
    tooltip.className = "word-tooltip";
    tooltip.id = "word-tooltip";

    // ——— Başlık ———
    let html = `<div class="tooltip-header">
<strong>${word}</strong>
<span class="tooltip-tr">🇹🇷 ${meanings.length} anlam</span>
</div>`;

    // ——— Her Türkçe anlam + İngilizce örnek cümle + Türkçe çevirisi ———
    html += '<div class="tooltip-meanings">';
    html += '<div class="meaning-group">';
    html += '<div class="meaning-pos">📖 TÜRKÇE ANLAMLARI & ÖRNEKLER</div>';
    meanings.forEach((m, i) => {
        html += `<div class="meaning-item">
<div class="meaning-def">${i + 1}. ${m}</div>
<div class="meaning-example" data-exslot="${i}">⏳ örnek hazırlanıyor...</div>
<div class="meaning-example-tr" data-exslot-tr="${i}"></div>
</div>`;
    });
    html += '</div>'; // meaning-group

    // ——— Sözlükten diğer Türkçe anlamlar (tür bazında) ———
    html += `<div class="meaning-group contextual-group">
<div class="meaning-pos">🇹🇷 SÖZLÜKTEN DİĞER ANLAMLAR</div>
<div class="syn-list" data-dictmeanings>⏳ aranıyor...</div>
</div>`;

    // ——— Eş anlamlılar (ikame kelimeler) ———
    html += `<div class="meaning-group contextual-group">
<div class="meaning-pos">🔁 EŞ ANLAMLILAR (İKAME)</div>
<div class="syn-list" data-synlist>⏳ aranıyor...</div>
</div>`;

    html += '</div>'; // tooltip-meanings

    tooltip.innerHTML = html;
    document.body.appendChild(tooltip);
    positionTooltip(tooltip, spanElement);

    // Tooltip fare olayları
    tooltip.addEventListener('mouseenter', () => clearTimeout(hideTooltipTimer));
    tooltip.addEventListener('mouseleave', () => scheduleHideTooltip(300));

    // Bu balon artık açık → aç/kapa için kaydet
    activeTooltipEl = spanElement;

    // Sözlük anlamları tek istektir, hemen başlat;
    // örnekler ve eş anlamlılar sırayla doldurulur (hız sınırını aşmamak için)
    fillKeywordDictMeanings(word, tooltip);
    fillKeywordExamples(word, meanings, tooltip)
        .then(() => fillKeywordSynonyms(word, tooltip));
}

// ——— Anahtar kelimenin sözlükteki tüm Türkçe anlamlarını doldur ———
async function fillKeywordDictMeanings(word, tooltip) {
    const container = tooltip.querySelector('[data-dictmeanings]');
    if (!container) return;

    const d = await fetchTurkishDict(word);
    if (!document.body.contains(tooltip)) return;

    const lines = [];
    d.groups.forEach(g => {
        lines.push(`<div class="syn-item"><span class="syn-en">${translatePosLabel(g.pos)}:</span> <span class="syn-tr">${g.terms.join(", ")}</span></div>`);
    });
    if (lines.length === 0) {
        const flat = [d.main].concat(d.alts).filter(Boolean);
        if (flat.length > 0) {
            lines.push(`<div class="syn-item"><span class="syn-tr">${flat.join(", ")}</span></div>`);
        }
    }
    container.innerHTML = lines.length > 0
        ? lines.join('')
        : '— bulunamadı — (tekrar dokunarak yeniden deneyin)';
}

// =========================================================
// EŞ ANLAMLILAR (İKAME KELİMELER) — Free Dictionary API'den
// =========================================================
async function fillKeywordSynonyms(word, tooltip) {
    const container = tooltip.querySelector('[data-synlist]');
    if (!container) return;

    const lower = word.toLowerCase().replace(/[^a-z']/g, '');
    let synonyms = [];

    if (lower.length >= 2) {
        try {
            const defs = await getDictionaryDefinitions(lower, word);
            const set = new Set();
            for (const meaning of (defs || [])) {
                (meaning.synonyms || []).forEach(s => set.add(s));
                for (const def of (meaning.definitions || [])) {
                    (def.synonyms || []).forEach(s => set.add(s));
                }
            }
            synonyms = [...set]
                .filter(s => s && s.toLowerCase() !== lower)
                .slice(0, 5);
        } catch (_) {}
    }

    if (!document.body.contains(tooltip)) return;

    if (synonyms.length === 0) {
        container.textContent = '— eş anlamlı bulunamadı —';
        return;
    }

    // İngilizce eş anlamlıları listele, Türkçelerini sırayla ekle
    container.innerHTML = synonyms.map((s, i) =>
        `<div class="syn-item"><span class="syn-en">${s}</span> <span class="syn-tr" data-syntr="${i}"></span></div>`
    ).join('');

    for (let i = 0; i < synonyms.length; i++) {
        if (!document.body.contains(tooltip)) return;
        const trEl = container.querySelector(`[data-syntr="${i}"]`);
        if (!trEl) continue;
        const tr = await translateToTurkish(synonyms[i]);
        if (tr) trEl.textContent = `(${tr})`;
    }
}

// =========================================================
// HER ANLAM İÇİN İNGİLİZCE ÖRNEK CÜMLE + TÜRKÇE ÇEVİRİSİ
// =========================================================
async function fillKeywordExamples(word, meanings, tooltip) {
    const cardData = cards[currentIndex];
    const lower = word.toLowerCase().replace(/[^a-z']/g, '');

    // ——— Örnek havuzu oluştur: { en, tr } ———
    const pool = [];

    // 1. Mevcut kartın cümlesi (gerçek bağlam, Türkçesi hazır)
    if (cardData.sentence) {
        pool.push({ en: cardData.sentence, tr: cardData.turkish || '' });
    }

    // 2. Aynı kelimeyi içeren diğer kartlar (Türkçeleri hazır)
    for (let i = 0; i < cards.length && pool.length < meanings.length + 2; i++) {
        if (i === currentIndex) continue;
        const s = cards[i].sentence || cards[i].english || '';
        const tokens = s.toLowerCase().match(/[\w']+/g) || [];
        if (tokens.includes(lower)) {
            pool.push({ en: s, tr: cards[i].turkish || '' });
        }
    }

    // 3. Yeterli örnek yoksa sözlük API'sinden örnek cümleler çek (Türkçesi sonra çevrilir)
    if (pool.length < meanings.length && lower.length >= 2) {
        try {
            const defs = await getDictionaryDefinitions(lower, word);
            for (const meaning of (defs || [])) {
                for (const def of (meaning.definitions || [])) {
                    if (def.example) {
                        pool.push({ en: def.example, tr: null });
                        if (pool.length >= meanings.length + 2) break;
                    }
                }
                if (pool.length >= meanings.length + 2) break;
            }
        } catch (_) {}
    }

    // ——— Her anlama bir örnek ata ve gerekirse Türkçeye çevir ———
    for (let i = 0; i < meanings.length; i++) {
        const enEl = tooltip.querySelector(`[data-exslot="${i}"]`);
        const trEl = tooltip.querySelector(`[data-exslot-tr="${i}"]`);
        if (!enEl) continue;

        const ex = pool[i];
        if (!ex) {
            enEl.textContent = '— örnek bulunamadı —';
            if (trEl) trEl.style.display = 'none';
            continue;
        }

        enEl.textContent = `💬 ${ex.en}`;
        let tr = ex.tr;
        if (!tr) {
            if (trEl) trEl.textContent = '🔄 Türkçesi çevriliyor...';
            tr = await translateToTurkish(ex.en);
        }
        if (trEl) {
            if (tr) {
                trEl.textContent = `🇹🇷 ${tr}`;
            } else {
                trEl.style.display = 'none';
            }
        }
    }
}

// =========================================================
// DİĞER KELİMELER İÇİN DICTIONARY API ÇEVİRİSİ
// =========================================================
async function showWordMeanings(lower, originalWord, spanElement) {
    // Balonu BEKLEMEDEN hemen aç: kullanıcı tıklamasının karşılığını anında görür,
    // çeviri ve tanımlar geldikçe içerik dolar.
    showLoadingTooltip(spanElement, originalWord);

    const [trDict, meanings] = await Promise.all([
        fetchTurkishDict(originalWord),
        getDictionaryDefinitions(lower, originalWord)
    ]);

    // Kullanıcı bu arada balonu kapattıysa ya da başka kelimeye geçtiyse dokunma
    if (activeTooltipEl !== spanElement || !document.getElementById("word-tooltip")) return;
    showRichTooltip(spanElement, originalWord, trDict.main, meanings, trDict);
}

// ——— İngilizce sözcük türü etiketini Türkçeleştir ———
function translatePosLabel(pos) {
    return String(pos || "")
        .replace('verb', 'fiil')
        .replace('noun', 'isim')
        .replace('adjective', 'sıfat')
        .replace('adverb', 'zarf')
        .replace('preposition', 'edat')
        .replace('conjunction', 'bağlaç')
        .replace('pronoun', 'zamir')
        .replace('interjection', 'ünlem')
        .replace('determiner', 'belirteç')
        .replace('exclamation', 'ünlem')
        .replace('abbreviation', 'kısaltma')
        .replace('phrase', 'kalıp');
}

// ——— Asılı kalan istekleri kesen fetch (ölü servis tüm uygulamayı kilitlemesin) ———
function fetchWithTimeout(url, options, timeoutMs) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs || 6000);
    const opts = Object.assign({}, options || {}, { signal: controller.signal });
    return fetch(url, opts).finally(() => clearTimeout(timer));
}

// ——— Sonuç doğrulama: "-", "…" gibi harf içermeyen "çeviriler" geçersizdir ———
function hasLetters(s) {
    return /[a-zA-ZçğıöşüÇĞİÖŞÜâîû]/.test(s || "");
}

// ——— Çeviri sağlayıcıları (sırayla denenir) ———
async function translateViaGoogle(text) {
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=tr&dt=t&q=${encodeURIComponent(text)}`;
    const res = await fetchWithTimeout(url, null, 6000);
    if (!res.ok) throw new Error("google http " + res.status);
    const data = await res.json();
    const translation = (data[0] || []).map(seg => seg && seg[0]).join("").trim();
    if (!hasLetters(translation)) throw new Error("google empty");
    return translation;
}

async function translateViaMyMemory(text) {
    const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=en|tr`;
    const res = await fetchWithTimeout(url, null, 8000);
    if (!res.ok) throw new Error("mymemory http " + res.status);
    const data = await res.json();
    const translation = (data.responseStatus === 200
        ? (data.responseData && data.responseData.translatedText) || ""
        : "").trim();
    if (!hasLetters(translation)) throw new Error("mymemory empty");
    return translation;
}

// =========================================================
// TÜRKÇE SÖZLÜK KARŞILIKLARI — tam anlam + birden çok anlam
// Google'ın sözlük verisi (dt=bd) tür bazında anlam listesi verir;
// boş kalırsa MyMemory'nin alternatif eşleşmeleri kullanılır.
// Dönen değer: { main: "ana karşılık", groups: [{pos, terms[]}], alts: [] }
// =========================================================
async function fetchTurkishDict(text) {
    const key = String(text || "").toLowerCase().trim();
    if (!key) return { main: "", groups: [], alts: [] };

    // Gömülü kalıp sözlüğü: internetsiz, anında ve her zaman çalışır
    if (PHRASE_TR[key]) {
        return { main: PHRASE_TR[key], groups: [], alts: [] };
    }

    const cached = dictCache[key];
    if (cached && cached.data) return cached.data;
    if (cached && cached.promise) return cached.promise;

    const promise = (async () => {
        let main = "";
        const groups = [];
        const alts = [];

        // 1) Google: ana çeviri + tür bazında anlam listesi
        try {
            const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=tr&dt=t&dt=bd&q=${encodeURIComponent(text)}`;
            const res = await fetchWithTimeout(url, null, 6000);
            if (res.ok) {
                const data = await res.json();
                const m = (data[0] || []).map(seg => seg && seg[0]).join("").trim();
                if (hasLetters(m)) main = m;
                if (Array.isArray(data[1])) {
                    for (const g of data[1]) {
                        if (g && typeof g[0] === "string" && Array.isArray(g[1])) {
                            const terms = g[1]
                                .filter(t => typeof t === "string" && hasLetters(t))
                                .slice(0, 6);
                            if (terms.length) groups.push({ pos: g[0], terms });
                        }
                    }
                }
            }
        } catch (_) {}

        // 2) Ana çeviri ya da anlam listesi eksikse MyMemory'den tamamla
        if (!main || groups.length === 0) {
            try {
                const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=en|tr`;
                const res = await fetchWithTimeout(url, null, 8000);
                if (res.ok) {
                    const data = await res.json();
                    if (data.responseStatus === 200) {
                        const m = ((data.responseData || {}).translatedText || "").trim();
                        if (!main && hasLetters(m)) main = m;
                        const seen = new Set([main.toLowerCase()]);
                        for (const match of (data.matches || [])) {
                            const t = (match && typeof match.translation === "string" ? match.translation : "").trim();
                            const tl = t.toLowerCase();
                            if (hasLetters(t) && !seen.has(tl) && t.length <= 60) {
                                seen.add(tl);
                                alts.push(t);
                            }
                            if (alts.length >= 5) break;
                        }
                    }
                }
            } catch (_) {}
        }

        const result = { main, groups, alts };
        if (main || groups.length || alts.length) {
            dictCache[key] = { data: result };
            if (main && !phraseCache[key]) phraseCache[key] = main; // cümle önbelleğiyle paylaş
            saveTrCache();
        } else {
            delete dictCache[key]; // başarısız → sonraki dokunuşta yeniden dene
        }
        return result;
    })();

    dictCache[key] = { promise };
    return promise;
}

// ——— Sözlükteki ek karşılıkları tek listeye indir (ana karşılık hariç) ———
function dictExtraTerms(d) {
    const seen = new Set([(d.main || "").toLowerCase()]);
    const out = [];
    const push = t => {
        const tl = String(t).toLowerCase();
        if (!seen.has(tl)) { seen.add(tl); out.push(t); }
    };
    d.groups.forEach(g => g.terms.forEach(push));
    d.alts.forEach(push);
    return out.slice(0, 8);
}

// ——— Cümle / ifade çevirisi ———
// Başarısız denemeler önbelleğe YAZILMAZ: bir sonraki tıklamada yeniden denenir.
// Aynı metin için uçuştaki istek paylaşılır (promise önbelleği).
async function translateToTurkish(text) {
    const key = String(text || "").toLowerCase().trim();
    if (!key) return "";
    if (PHRASE_TR[key]) return PHRASE_TR[key]; // gömülü kalıp sözlüğü
    const cached = phraseCache[key];
    if (typeof cached === "string") return cached;
    if (cached && cached.promise) return cached.promise;

    const promise = (async () => {
        try {
            const tr = await translateViaGoogle(text);
            phraseCache[key] = tr;
            saveTrCache();
            return tr;
        } catch (_) {}
        try {
            const tr = await translateViaMyMemory(text);
            phraseCache[key] = tr;
            saveTrCache();
            return tr;
        } catch (_) {}
        delete phraseCache[key];
        return "";
    })();

    phraseCache[key] = { promise };
    return promise;
}

// ——— Sözlük tanımları (Free Dictionary API) ———
// Sadece başarılı yanıtlar (ve gerçek "kelime yok" = 404) önbelleğe alınır;
// ağ hatası / hız sınırı sonraki denemede yeniden sorgulanır.
async function getDictionaryDefinitions(lower, originalWord) {
    if (defCache[lower]) {
        if (defCache[lower].data !== undefined) return defCache[lower].data;
        return defCache[lower].promise;
    }

    const promise = (async () => {
        try {
            const response = await fetchWithTimeout(
                `https://api.dictionaryapi.dev/api/v2/entries/en/${encodeURIComponent(originalWord)}`,
                null, 7000
            );
            if (response.ok) {
                const data = await response.json();
                const meanings = data[0]?.meanings || [];
                defCache[lower] = { data: meanings, promise: null };
                return meanings;
            }
            if (response.status === 404) {
                defCache[lower] = { data: [], promise: null };
                return [];
            }
        } catch (_) {}
        delete defCache[lower];
        return [];
    })();

    defCache[lower] = { data: undefined, promise };
    return promise;
}

// =========================================================
// ZENGİN TOOLTIP (diğer kelimeler için)
// İngilizce tanımlar + her örnek cümlenin Türkçe çevirisi
// =========================================================
function showRichTooltip(spanElement, word, translation, meanings, trDict) {
    hideTooltip();

    const tooltip = document.createElement("div");
    tooltip.className = "word-tooltip";
    tooltip.id = "word-tooltip";

    let headerHtml = `<div class="tooltip-header">
<strong>${word}</strong>`;
    if (translation && translation !== "🔄" && translation !== "❌") {
        headerHtml += `<span class="tooltip-tr">🇹🇷 ${translation}</span>`;
    } else if (translation === "🔄") {
        headerHtml += `<span class="tooltip-tr loading">🔄 çeviriliyor...</span>`;
    }
    headerHtml += `</div>`;

    let exIndex = 0;
    const examplesToTranslate = [];
    let defIndex = 0;
    const defsToTranslate = [];

    let meaningsHtml = '<div class="tooltip-meanings">';

    // ——— EN ÜSTTE: Türkçe anlamlar (tür bazında, birden çok karşılık) ———
    if (trDict && (trDict.groups.length > 0 || trDict.alts.length > 0)) {
        meaningsHtml += `<div class="meaning-group">`;
        meaningsHtml += `<div class="meaning-pos">🇹🇷 TÜRKÇE ANLAMLARI</div>`;
        if (trDict.groups.length > 0) {
            trDict.groups.forEach(g => {
                meaningsHtml += `<div class="meaning-item">
<div class="meaning-def"><b>${translatePosLabel(g.pos)}:</b> ${g.terms.join(", ")}</div>
</div>`;
            });
        } else {
            meaningsHtml += `<div class="meaning-item">
<div class="meaning-def">${[trDict.main].concat(trDict.alts).filter(Boolean).join(", ")}</div>
</div>`;
        }
        meaningsHtml += `</div>`;
    }

    if (meanings && meanings.length > 0) {
        meanings.forEach((meaning) => {
            const posLabel = translatePosLabel(meaning.partOfSpeech);

            meaningsHtml += `<div class="meaning-group">`;
            meaningsHtml += `<div class="meaning-pos">${posLabel}</div>`;

            meaning.definitions.forEach((def, di) => {
                meaningsHtml += `<div class="meaning-item">`;
                meaningsHtml += `<div class="meaning-def">${di + 1}. ${def.definition}</div>`;
                // İlk birkaç tanımın Türkçe karşılığı çevrilir; hepsini çevirmek
                // ücretsiz çeviri servisinin hız sınırına takılmaya yol açıyordu
                if (defIndex < 4) {
                    const didx = defIndex++;
                    defsToTranslate.push({ idx: didx, text: def.definition });
                    meaningsHtml += `<div class="meaning-def-tr" data-deftr="${didx}">🔄 Türkçesi çevriliyor...</div>`;
                }
                if (def.example) {
                    meaningsHtml += `<div class="meaning-example">💬 ${def.example}</div>`;
                    if (exIndex < 3) {
                        const idx = exIndex++;
                        examplesToTranslate.push({ idx: idx, text: def.example });
                        meaningsHtml += `<div class="meaning-example-tr" data-extr="${idx}">🔄 Türkçesi çevriliyor...</div>`;
                    }
                    meaningsHtml += `<button class="example-btn" data-example="${encodeURIComponent(def.example)}" data-word="${word}">▶ Kartta Göster</button>`;
                }
                meaningsHtml += `</div>`;
            });

            meaningsHtml += `</div>`;
        });
    } else if (!trDict || (trDict.groups.length === 0 && trDict.alts.length === 0)) {
        meaningsHtml += `<div class="meaning-item">
<div class="meaning-def">${translation || "Çeviriye ulaşılamadı — kelimeye tekrar dokunarak yeniden deneyin"}</div>
</div>`;
    }
    meaningsHtml += '</div>';

    tooltip.innerHTML = headerHtml + meaningsHtml;
    document.body.appendChild(tooltip);
    positionTooltip(tooltip, spanElement);

    // Tanımların ve örneklerin Türkçesini SIRAYLA çevir (paralel istekler ücretsiz
    // çeviri servisinde hız sınırına takıldığı için teker teker yapılır)
    (async () => {
        const tasks = [
            ...defsToTranslate.map(d => ({
                sel: `[data-deftr="${d.idx}"]`,
                text: d.text,
                fmt: tr => `(🇹🇷 ${tr})`
            })),
            ...examplesToTranslate.map(ex => ({
                sel: `[data-extr="${ex.idx}"]`,
                text: ex.text,
                fmt: tr => `🇹🇷 ${tr}`
            }))
        ];

        for (const t of tasks) {
            // Balon kapandıysa boşuna çeviri yapma
            if (!document.body.contains(tooltip)) return;
            const el = tooltip.querySelector(t.sel);
            if (!el) continue;
            const tr = await translateToTurkish(t.text);
            if (tr) {
                el.textContent = t.fmt(tr);
            } else {
                el.style.display = "none";
            }
        }
    })();

    tooltip.querySelectorAll('.example-btn').forEach(btn => {
        btn.addEventListener('click', function (e) {
            e.stopPropagation();
            const example = decodeURIComponent(this.dataset.example);
            const word = this.dataset.word;
            const turkish = this.dataset.turkish
                ? decodeURIComponent(this.dataset.turkish)
                : '';
            const ctxIndex = this.dataset.ctxindex !== undefined
                ? parseInt(this.dataset.ctxindex)
                : null;
            showExampleInCard(example, word, turkish, ctxIndex);
            hideTooltip();
        });
    });

    tooltip.addEventListener('mouseenter', () => clearTimeout(hideTooltipTimer));
    tooltip.addEventListener('mouseleave', () => scheduleHideTooltip(300));

    // Bu balon artık açık → aç/kapa için kaydet
    activeTooltipEl = spanElement;
}

// =========================================================
// ÖRNEK CÜMLEYİ KARTTA GÖSTER
// =========================================================
function showExampleInCard(exampleSentence, word, turkishTranslation, ctxIndex) {
    window.speechSynthesis.cancel();
    isShowingExample = true;
    originalCardIndex = currentIndex;
    selAnchorIdx = null;
    const card = document.getElementById("card");

    // Kalıplar tek parça, geri kalanı kelime kelime; eşleşen kelime vurgulanır
    const html = buildHtmlWithPhrases(exampleSentence, (t) => tokenizeSegmentHighlight(t, word));

    card.innerHTML = `
<div class="example-card-content">
<div class="example-badge">📝 Örnek Cümle</div>
<h2>${html}</h2>
${turkishTranslation ? `<div class="example-translation">🇹🇷 ${turkishTranslation}</div>` : ''}
<button class="back-to-card-btn">← Ana Karta Dön</button>
</div>
`;

    attachWordListeners(card);

    card.querySelector('.back-to-card-btn').addEventListener('click', function (e) {
        e.stopPropagation();
        showOriginalCard();
    });

    updateCounter();
    speakEnglish(exampleSentence);
}

function showOriginalCard() {
    isShowingExample = false;
    currentIndex = originalCardIndex;
    showEnglish();
}

function hideTooltip() {
    const existing = document.getElementById("word-tooltip");
    if (existing) {
        existing.remove();
    }
    activeTooltipEl = null;
}

// =========================================================
// KART TIKLAMA – ÇEVİRİYİ GÖSTER / GİZLE
// =========================================================
document.getElementById("card").addEventListener("click", function () {
    if (isShowingExample) {
        showOriginalCard();
        return;
    }
    const tooltip = document.getElementById("word-tooltip");
    if (tooltip) {
        hideTooltip();
        clearSelection();
        return;
    }
    // Seçim modunda kartın boş alanına dokunmak kartı çevirmez;
    // sadece yarım kalan seçimi temizler (yanlışlıkla çevrilme önlenir)
    if (selectMode) {
        clearSelection();
        return;
    }
    const card = this;
    card.classList.add("flip-effect");
    setTimeout(() => {
        if (showingTurkish) {
            showEnglish();
        } else {
            showTurkish();
        }
        card.classList.remove("flip-effect");
    }, 300);
});

// Tooltip dışına tıklandığında tooltip'i kapat
document.addEventListener("click", function (e) {
    const tooltip = document.getElementById("word-tooltip");
    if (tooltip && !tooltip.contains(e.target) && !e.target.closest('.word-clickable') && !e.target.closest('.word-keyword') && !e.target.closest('.word-phrase')) {
        hideTooltip();
        clearSelection();
    }
});

// =========================================================
// ÖNCEKİ / SONRAKİ KART + SESLENDİRME BUTONU
// =========================================================
document.getElementById("next-btn")?.addEventListener("click", nextCard);
document.getElementById("prev-btn")?.addEventListener("click", prevCard);
document.getElementById("speak-btn")?.addEventListener("click", function (e) {
    e.stopPropagation();
    speakCurrentSentence();
});
document.getElementById("copy-btn")?.addEventListener("click", function (e) {
    e.stopPropagation();
    copyCurrentSentence();
});

// ——— 🖍 Seçim modu aç/kapa ———
document.getElementById("select-btn")?.addEventListener("click", function (e) {
    e.stopPropagation();
    selectMode = !selectMode;
    this.classList.toggle("select-active", selectMode);
    document.getElementById("card").classList.toggle("select-mode", selectMode);
    hideTooltip();
    clearSelection();
    showToast(selectMode
        ? "🖍 Seçim modu açık: önce ilk kelimeye, sonra son kelimeye dokun"
        : "Seçim modu kapatıldı");
});

// =========================================================
// CÜMLEYİ KOPYALA (📋 butonu) — İngilizce cümlenin tamamını panoya kopyala
// =========================================================
function copyCurrentSentence() {
    const text = getCurrentSentenceText().trim();
    if (!text) return;

    const btn = document.getElementById("copy-btn");
    const flash = (symbol) => {
        if (!btn) return;
        btn.textContent = symbol;
        setTimeout(() => { btn.textContent = "📋"; }, 1200);
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text)
            .then(() => flash("✅"))
            .catch(() => flash(fallbackCopy(text) ? "✅" : "❌"));
    } else {
        flash(fallbackCopy(text) ? "✅" : "❌");
    }
}

// Pano API'si çalışmazsa (eski tarayıcı / izin yok) yedek kopyalama
function fallbackCopy(text) {
    try {
        const ta = document.createElement("textarea");
        ta.value = text;
        ta.style.position = "fixed";
        ta.style.top = "-1000px";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        const ok = document.execCommand("copy");
        document.body.removeChild(ta);
        return ok;
    } catch (_) {
        return false;
    }
}

function goToCard(index) {
    window.speechSynthesis.cancel();
    hideTooltip();
    clearTimers();
    if (isShowingExample) {
        isShowingExample = false;
    }
    const card = document.getElementById("card");
    card.classList.add("fade-out");
    setTimeout(() => {
        currentIndex = index;
        showEnglish();
        card.classList.remove("fade-out");
    }, 300);
}

function nextCard() {
    window.speechSynthesis.cancel();
    hideTooltip();
    clearTimers();
    if (isShowingExample) {
        currentIndex = originalCardIndex;
        isShowingExample = false;
    }
    const card = document.getElementById("card");
    card.classList.add("fade-out");
    setTimeout(() => {
        currentIndex++;
        if (currentIndex >= cards.length) {
            currentIndex = 0;
        }
        showEnglish();
        card.classList.remove("fade-out");
    }, 300);
}

function prevCard() {
    window.speechSynthesis.cancel();
    hideTooltip();
    clearTimers();
    if (isShowingExample) {
        currentIndex = originalCardIndex;
        isShowingExample = false;
    }
    const card = document.getElementById("card");
    card.classList.add("fade-out");
    setTimeout(() => {
        currentIndex--;
        if (currentIndex < 0) {
            currentIndex = cards.length - 1;
        }
        showEnglish();
        card.classList.remove("fade-out");
    }, 300);
}

// =========================================================
// KLAVYE KISAYOLLARI
// =========================================================
document.addEventListener("keydown", function (e) {
    const pageInput = document.getElementById("page-input");
    if (document.activeElement === pageInput) {
        if (e.key === "Enter") {
            const num = parseInt(pageInput.value);
            if (num >= 1 && num <= cards.length) {
                goToCard(num - 1);
            }
            pageInput.value = "";
        }
        return;
    }

    if (e.key === "ArrowLeft") {
        e.preventDefault();
        prevCard();
    } else if (e.key === "ArrowRight") {
        e.preventDefault();
        nextCard();
    }
});

// =========================================================
// SAYFA ATLAMA INPUT
// =========================================================
document.getElementById("page-input")?.addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
        const num = parseInt(this.value);
        if (num >= 1 && num <= cards.length) {
            goToCard(num - 1);
        }
        this.value = "";
    }
});

document.getElementById("page-input")?.addEventListener("blur", function () {
    this.value = "";
});

// =========================================================
// DOKUNMATİK KAYDIRMA (SWIPE) — parmakla sağa/sola geçiş
// Hem dikey hem yatay kullanımda çalışır.
// =========================================================
(function setupSwipeNavigation() {
    const card = document.getElementById("card");
    if (!card) return;

    let startX = 0, startY = 0, startTime = 0, tracking = false;
    const MIN_DISTANCE = 50;    // en az yatay kaydırma (px)
    const OFF_AXIS_RATIO = 0.8; // dikey hareket bu orandan büyükse swipe sayılmaz
    const MAX_DURATION = 1000;  // ms

    // Swipe sonrası tarayıcının ürettiği "click"i yut:
    // kartın çevrilmesini / kelime balonunun açılmasını engeller.
    function swallowNextClick() {
        const handler = function (e) {
            e.stopPropagation();
            e.preventDefault();
            cleanup();
        };
        const cleanup = function () {
            document.removeEventListener("click", handler, true);
            clearTimeout(timer);
        };
        document.addEventListener("click", handler, true);
        const timer = setTimeout(cleanup, 700);
    }

    card.addEventListener("touchstart", function (e) {
        if (e.touches.length !== 1) { tracking = false; return; }
        const t = e.touches[0];
        startX = t.clientX;
        startY = t.clientY;
        startTime = Date.now();
        tracking = true;
    }, { passive: true });

    card.addEventListener("touchend", function (e) {
        if (!tracking) return;
        tracking = false;

        const t = e.changedTouches[0];
        const dx = t.clientX - startX;
        const dy = t.clientY - startY;
        const dt = Date.now() - startTime;

        if (dt > MAX_DURATION) return;
        if (Math.abs(dx) < MIN_DISTANCE) return;
        if (Math.abs(dy) > Math.abs(dx) * OFF_AXIS_RATIO) return; // dikey kaydırma → yok say

        swallowNextClick();

        if (dx < 0) {
            nextCard();   // sola kaydır → sonraki kart
        } else {
            prevCard();   // sağa kaydır → önceki kart
        }
    }, { passive: true });

    card.addEventListener("touchcancel", function () {
        tracking = false;
    }, { passive: true });
}());

// =========================================================
// BAŞLAT
// =========================================================
loadCards();
