import streamlit as st
import pandas as pd
import os
import plotly.express as px

# Excel dosya adları
KELIME_FILE = "kelimeler.xlsx"
USER_FILE = "kullanicilar.xlsx"

# --- TÜM 320 KELİMENİN EKSİKSİZ OTOMATİK YÜKLEME KODU ---
if not os.path.exists(KELIME_FILE) or (os.path.exists(KELIME_FILE) and len(pd.read_excel(KELIME_FILE)) < 200):
    hazir_sozluk = {
        "Ingilizce": [
            "also", "mistake", "discover", "even", "decision", "seriously", "amongst", "between", "whole", "tale",
            "be", "cares", "brilliant", "criminal", "common", "communicate", "escalator", "imagine", "ever",
            "adventure",
            "expectations", "freak", "kind", "urban", "each other", "critics", "trade", "fell", "through", "fair",
            "almost", "excellent", "price", "cheap", "protect", "treatment", "discuss", "told", "confident", "reason",
            "pleasure", "grown man", "melt", "vanished", "move", "take care", "ridiculous", "thirsty", "imaginative",
            "wish",
            "skin", "hurt", "happened", "being", "facility", "field", "coast", "sort", "east", "west",
            "south", "north", "expect", "explain", "pain", "realize", "legacy", "dreadful", "in front of x",
            "plenty of",
            "look forward", "dirty", "distance", "awful", "neighbours", "adopted", "appreciate", "embarrass", "bargain",
            "assume",
            "upon", "flaw", "flawless", "us", "ignore", "blink", "deal", "hubris", "except", "recommend",
            "shy", "suggest", "delicate", "balance", "miserable", "hurry", "sweat", "colleague", "bald", "meet",
            "met", "case", "dare", "delegate", "disgust", "belong", "discount", "huge", "amount", "influence",
            "climb", "pour", "hang", "growth", "shrink", "maintain", "independent", "gather", "feature", "firm",
            "widespread", "accurate", "blame", "deny", "admit", "refuse", "complain", "encourage", "warn", "threaten",
            "request", "require", "demand", "supply", "provide", "increase", "decrease", "improve", "destroy", "create",
            "prevent", "allow", "avoid", "achieve", "succeed", "fail", "manage", "afford", "waste", "spend",
            "save", "borrow", "lend", "owe", "earn", "gain", "lose", "win", "beat", "defeat",
            "celebrate", "congratulate", "support", "oppose", "defend", "attack", "protect", "rescue", "survive",
            "destroy",
            "damage", "repair", "fix", "clean", "wash", "dirty", "tidy", "messy", "organized", "confused",
            "excited", "bored", "tired", "exhausted", "sleepy", "awake", "alive", "dead", "born", "die",
            "kill", "murder", "steal", "rob", "burgle", "arrest", "accuse", "suspect", "innocent", "guilty",
            "judge", "court", "law", "rule", "government", "politics", "election", "vote", "candidate", "president",
            "minister", "ambassador", "citizen", "foreigner", "stranger", "guest", "host", "customer", "client",
            "employee",
            "employer", "manager", "director", "staff", "crew", "passenger", "driver", "pilot", "captain", "soldier",
            "officer", "army", "navy", "air force", "war", "peace", "battle", "victory", "defeat", "weapon",
            "gun", "knife", "sword", "shield", "armor", "helmet", "bullet", "explosion", "bomb", "danger",
            "safety", "hazard", "risk", "accident", "injury", "wound", "blood", "bleed", "pain", "ache",
            "illness", "disease", "infection", "virus", "bacteria", "cure", "medicine", "pill", "drug", "hospital",
            "clinic", "doctor", "nurse", "patient", "surgery", "operation", "recovery", "health", "fitness", "exercise",
            "train", "practice", "match", "game", "sport", "team", "player", "coach", "referee", "stadium",
            "gym", "pool", "field", "court", "track", "race", "run", "jump", "throw", "catch",
            "hit", "kick", "score", "goal", "point", "win", "trophy", "medal", "champion", "tournament"
        ],
        "Turkce": [
            "ayrıca", "hata", "keşfetmek", "hatta", "karar", "cidden", "aramızda", "arasında", "tüm", "masal",
            "olmak", "umursamak", "muhteşem", "suçlu", "yaygın", "iletişim", "yürüyen merdiven", "hayal etmek", "asla",
            "macera",
            "beklenti", "çatlak ucube", "tip tür", "kentsel şehir", "birbirine göre", "eleştirmenler", "ticaret",
            "düşmüş", "baştan sona", "adil",
            "neredeyse", "harika", "fiyat", "ucuz", "korumak", "hasta tedavi", "tartışmak", "söylemek",
            "kendinden emin", "sebep",
            "zevk", "yetişkin", "eritmek", "ortadan kaybolmak", "hareket", "dikkatli ol", "gülünç", "susuz", "yaratıcı",
            "dilek",
            "deri", "acıtmak", "olmak", "varlık", "tesis", "alan", "sahil", "düzenlemek", "doğu", "batı",
            "güney", "kuzey", "beklemek", "açıklamak", "ağrı", "fark etmek", "miras", "korkunç", "x in önünde",
            "bol çok",
            "dört gözle beklemek", "kirli", "mesafe", "berbat", "komşu", "kabul edilen", "takdirle karşılamak",
            "utanmak", "pazarlık", "farz etmek",
            "üzerine", "kusur", "kusursuz", "biz", "görmezden gelmek", "göz kırpmak", "anlaşma", "kibir", "dışında",
            "tavsiye",
            "utangaç", "önermek", "hassas", "denge", "sefil", "acele etmek", "terlemek", "iş arkadaşı", "kel",
            "ilk buluşma",
            "ikinci buluşma", "dava", "cesaret", "temsilci", "iğrenç", "ait olma", "indirim", "devasa", "miktar",
            "etki",
            "tırmanmak", "dökmek", "asmak", "büyüme", "küçülmek", "sürdürmek", "bağımsız", "toplamak", "özellik",
            "firma",
            "yaygın", "doğru", "suçlamak", "inkar etmek", "itiraf etmek", "reddetmek", "şikayet etmek", "teşvik etmek",
            "uyarmak", "tehdit etmek",
            "rica etmek", "gerekmek", "talep etmek", "sağlamak", "sağlamak", "artırmak", "azaltmak", "geliştirmek",
            "yok etmek", "yaratmak",
            "önlemek", "izin vermek", "kaçınmak", "başarmak", "başarmak", "başarısız olmak", "yönetmek", "gücü yetmek",
            "boşa harcamak", "harcamak",
            "biriktirmek", "ödünç almak", "ödünç vermek", "borçlu olmak", "kazanmak", "kazanmak", "kaybetmek",
            "kazanmak", "yenmek", "yenmek",
            "kutlamak", "tebrik etmek", "desteklemek", "karşı çıkmak", "savunmak", "saldırmak", "korumak", "kurtarmak",
            "hayatta kalmak", "yok etmek",
            "zarar vermek", "tamir etmek", "tamir etmek", "temizlemek", "yıkamak", "kirli", "düzenli", "dağınık",
            "organize", "kafası karışmış",
            "heyecanlı", "sıkılmış", "yorgun", "tükenmiş", "uykulu", "uyanık", "canlı", "ölü", "doğmuş", "ölmek",
            "öldürmek", "cinayet", "çalmak", "soymak", "hırsızlık yapmak", "tutuklamak", "suçlamak", "şüphelenmek",
            "masum", "suçlu",
            "yargıç", "mahkeme", "kanun", "kural", "hükümet", "politika", "seçim", "oy vermek", "aday", "başkan",
            "bakan", "büyükelçi", "vatandaş", "yabancı", "yabancı", "misafir", "ev sahibi", "müşteri", "müşteri",
            "çalışan",
            "işveren", "müdür", "direktör", "personel", "mürettebat", "yolcu", "sürücü", "pilot", "kaptan", "asker",
            "memur", "ordu", "deniz kuvvetleri", "hava kuvvetleri", "savaş", "barış", "戰鬥", "zafer", "yenilgi",
            "silah",
            "silah", "bıçak", "kılıç", "kalkan", "zırh", "kask", "mermi", "patlama", "bomba", "tehlike",
            "güvenlik", "tehlike", "risk", "kaza", "yaralanma", "yara", "kan", "kanamak", "ağrı", "ağrı",
            "hastalık", "hastalık", "enfeksiyon", "virüs", "bakteri", "tedavi", "ilaç", "hap", "ilaç", "hastane",
            "klinik", "doktor", "hemşire", "hasta", "cerrahi", "ameliyat", "iyileşme", "sağlık", "fitness", "egzersiz",
            "tren", "pratik", "maç", "oyun", "spor", "takım", "oyuncu", "antrenör", "hakem", "stadyum",
            "spor salonu", "havuz", "alan", "saha", "pist", "yarış", "koşmak", "atamak", "fırlatmak", "yakalamak",
            "vurmak", "tekmelemek", "skor", "gol", "puan", "kazanmak", "kupa", "madalya", "şampiyon", "turnuva"
        ]
    }
    pd.DataFrame(hazir_sozluk).to_excel(KELIME_FILE, index=False)

if not os.path.exists(USER_FILE):
    pd.DataFrame(columns=["Kullanici", "Toplam_Soru", "Dogru_Sayisi"]).to_excel(USER_FILE, index=False)

# --- UYGULAMA BAŞLANGICI ---
st.set_page_config(page_title="Kelime Ezber", page_icon="📝", layout="centered")
st.title("📝 Kelime Ezber Uygulaması")

# --- KULLANICI GİRİŞ SİSTEMİ ---
if "kullanici" not in st.session_state:
    st.subheader("👤 Giriş Yap / Kayıt Ol")
    isim = st.text_input("İsminizi Giriniz:").strip()
    if st.button("Sisteme Giriş Yap"):
        if isim:
            st.session_state.kullanici = isim
            df_users = pd.read_excel(USER_FILE)
            if isim.lower() != "admin" and isim not in df_users["Kullanici"].values:
                yeni_user = pd.DataFrame([{"Kullanici": isim, "Toplam_Soru": 0, "Dogru_Sayisi": 0}])
                df_users = pd.concat([df_users, yeni_user], ignore_index=True)
                df_users.to_excel(USER_FILE, index=False)
            st.rerun()
        else:
            st.error("Lütfen bir isim girin.")
else:
    st.sidebar.write(f"👋 Hoş geldin, **{st.session_state.kullanici}**")
    if st.sidebar.button("Çıkış Yap"):
        keys_to_clear = ["kullanici", "test_basladi", "test_bitti", "sorular", "current_index", "skor", "deneme",
                         "cevaplandi"]
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

    df_genel = pd.read_excel(KELIME_FILE)
    toplam_kelime_sayisi = len(df_genel)

    is_admin = st.session_state.kullanici.lower() == "admin"
    sekme_listesi = ["🚀 Test Ol", "➕ Kelime Ekle", "📊 Geçmiş İstatistiklerim"]
    if is_admin:
        sekme_listesi.append("🛠️ Yönetici Paneli")

    sekmeler = st.tabs(sekme_listesi)

    # --- KELİME EKLEME SEKMESİ ---
    with sekmeler[1]:
        st.subheader("Kelime Havuzuna Ekle")
        st.info(f"📊 Havuzdaki Güncel Kelime Sayısı: {toplam_kelime_sayisi}")

        yeni_ing = st.text_input("İngilizce Kelime:", key="ing_ekle")
        yeni_tr = st.text_input("Türkçe Karşılığı:", key="tr_ekle")

        if st.button("Veritabanına Kaydet"):
            if yeni_ing and yeni_tr:
                df = pd.read_excel(KELIME_FILE)
                if not df.empty and (df['Ingilizce'].str.strip().str.lower() == yeni_ing.strip().lower()).any():
                    st.error(f"❌ '{yeni_ing}' kelimesi havuzda zaten mevcut!")
                else:
                    yeni_veri = pd.DataFrame([{"Ingilizce": yeni_ing.strip(), "Turkce": yeni_tr.strip()}])
                    df = pd.concat([df, yeni_veri], ignore_index=True)
                    df.to_excel(KELIME_FILE, index=False)
                    st.success(f"🎉 '{yeni_ing}' kelimesi başarıyla eklendi!")
                    st.rerun()
            else:
                st.error("Lütfen iki alanı da doldurun.")

    # --- GEÇMİŞ İSTATİSTİKLERİM SEKMESİ ---
    with sekmeler[2]:
        if is_admin:
            st.info(
                "Hacı, Admin hesabı için geçmiş istatistiği tutulmaz. Testleri normal kullanıcı adınla girip deneyebilirsin.")
        else:
            st.subheader(f"📊 {st.session_state.kullanici} - Genel Performans")
            df_users = pd.read_excel(USER_FILE)
            user_rows = df_users[df_users["Kullanici"] == st.session_state.kullanici]

            if user_rows.empty:
                st.info("Kullanıcı verisi bulunamadı.")
            else:
                user_row = user_rows.iloc[0]
                t_soru = user_row["Toplam_Soru"]
                t_dogru = user_row["Dogru_Sayisi"]
                t_yanlis = t_soru - t_dogru

                if t_soru == 0:
                    st.info("Henüz hiç test çözmemişsin hacı. Test çözdükçe burası şenlenecek.")
                else:
                    st.write(f"**Toplam Çözülen Soru:** {t_soru}")
                    st.write(f"**Toplam Doğru Cevap:** {t_dogru}")

                    basari_orani = int((t_dogru / t_soru) * 100)
                    fig_data = pd.DataFrame({"Durum": ["Doğru", "Yanlış"], "Sayı": [t_dogru, t_yanlis]})

                    fig = px.pie(fig_data, values="Sayı", names="Durum", hole=0.6, color="Durum",
                                 color_discrete_map={"Doğru": "#2ecc71", "Yanlış": "#e74c3c"})
                    fig.update_layout(
                        annotations=[
                            dict(text=f"%{basari_orani}<br>Doğruluk", x=0.5, y=0.5, font_size=24, showarrow=False)],
                        showlegend=True, height=350
                    )
                    st.plotly_chart(fig, use_container_width=True)

    # --- 🛠️ YÖNETİCİ PANELİ SEKMESİ ---
    if is_admin:
        with sekmeler[3]:
            st.subheader("🛠️ Kelime Silme ve Düzenleme Paneli")
            df_admin = pd.read_excel(KELIME_FILE)

            if df_admin.empty:
                st.info("Havuzda silinecek kelime yok hacı.")
            else:
                st.write(f"Toplam listelenen kelime: {len(df_admin)}")
                for idx, row in df_admin.iterrows():
                    col1, col2, col3 = st.columns([3, 3, 2])
                    col1.write(f"🇬🇧 {row['Ingilizce']}")
                    col2.write(f"🇹🇷 {row['Turkce']}")
                    if col3.button("Sil", key=f"del_{idx}"):
                        df_admin = df_admin.drop(idx)
                        df_admin.to_excel(KELIME_FILE, index=False)
                        st.success("Kelime başarıyla silindi!")
                        st.rerun()

    # --- TEST OLMA SEKMESİ ---
    with sekmeler[0]:
        df = pd.read_excel(KELIME_FILE)

        if df.empty:
            st.info("Henüz kelime eklenmemiş. Yan sekmeden kelime ekleyerek başlayın.")
        else:
            if "test_basladi" not in st.session_state:
                st.session_state.test_basladi = False
            if "test_bitti" not in st.session_state:
                st.session_state.test_bitti = False

            # --- AŞAMA 1: AYARLAR EKRANI ---
            if not st.session_state.test_basladi and not st.session_state.test_bitti:
                st.subheader("Test Ayarları")
                st.write(f"📊 **Toplam Kelime Havuzun:** {toplam_kelime_sayisi} kelime")

                yon = st.radio("Soru Yönü:", ["İngilizce -> Türkçe", "Türkçe -> İngilizce"])
                soru_sayisi = st.number_input("Kaç soru çözmek istersin?", min_value=1, max_value=toplam_kelime_sayisi,
                                              value=min(10, toplam_kelime_sayisi))

                if st.button("Testi Başlat"):
                    st.session_state.sorular = df.sample(n=int(soru_sayisi)).to_dict(orient="records")
                    st.session_state.yon = yon
                    st.session_state.current_index = 0
                    st.session_state.skor = 0
                    st.session_state.deneme = 0
                    st.session_state.test_basladi = True
                    st.session_state.test_bitti = False
                    st.session_state.cevaplandi = False
                    st.rerun()

            # --- AŞAMA 2: SORU EKRANI ---
            elif st.session_state.test_basladi and not st.session_state.test_bitti:
                idx = st.session_state.current_index
                total = len(st.session_state.sorular)
                row = st.session_state.sorular[idx]
                yon = st.session_state.yon

                soru_kelimesi = row["Ingilizce"] if yon == "İngilizce -> Türkçe" else row["Turkce"]
                dogru_cevap = row["Turkce"] if yon == "İngilizce -> Türkçe" else row["Ingilizce"]

                st.subheader(f"Soru {idx + 1} / {total}")
                st.progress((idx) / total)
                st.write(f"Kelime: ### **{soru_kelimesi}**")

                if st.session_state.deneme == 1 and not st.session_state.cevaplandi:
                    st.warning("⚠️ İlk tahminin yanlıştı, son 1 hakkın kaldı! Tekrar dene.")

                with st.form(key=f"form_{idx}_{st.session_state.deneme}", clear_on_submit=False):
                    user_cevap = st.text_input("Cevabınız:", key=f"input_{idx}_{st.session_state.deneme}",
                                               disabled=st.session_state.cevaplandi)

                    if not st.session_state.cevaplandi:
                        submit_button = st.form_submit_button(label="Cevapla (Enter)")
                        if submit_button:
                            if user_cevap.strip().lower() == dogru_cevap.strip().lower():
                                st.session_state.cevaplandi = True
                                st.session_state.is_correct = True
                                st.session_state.skor += 1
                                st.rerun()
                            else:
                                if st.session_state.deneme == 0:
                                    st.session_state.deneme = 1
                                    st.rerun()
                                else:
                                    st.session_state.cevaplandi = True
                                    st.session_state.is_correct = False
                                    st.rerun()
                    else:
                        if st.session_state.is_correct:
                            st.success("✅ Doğru!")
                            st.markdown(
                                '<audio src="https://assets.mixkit.co/active_storage/sfx/2568/2568-84.wav" autoplay></audio>',
                                unsafe_allow_html=True)
                        else:
                            st.error(f"❌ Yanlış! Doğru Cevap: **{dogru_cevap}**")

                        next_button = st.form_submit_button(
                            label="Sonraki Soru (Enter)" if idx + 1 < total else "Testi Bitir (Enter)")
                        if next_button:
                            if idx + 1 < total:
                                st.session_state.current_index += 1
                                st.session_state.cevaplandi = False
                                st.session_state.deneme = 0
                            else:
                                if st.session_state.kullanici.lower() != "admin":
                                    df_users = pd.read_excel(USER_FILE)
                                    user_idx = df_users[df_users["Kullanici"] == st.session_state.kullanici].index[0]
                                    df_users.at[user_idx, "Toplam_Soru"] += total
                                    df_users.at[user_idx, "Dogru_Sayisi"] += st.session_state.skor
                                    df_users.to_excel(USER_FILE, index=False)

                                st.session_state.test_basladi = False
                                st.session_state.test_bitti = True
                            st.rerun()

            # --- AŞAMA 3: SONUÇ EKRANI ---
            elif st.session_state.test_bitti:
                st.subheader("🎉 Test Sonucu")
                total = len(st.session_state.sorular)
                skor = st.session_state.skor

                st.metric(label="Bu Testteki Doğru Sayısı", value=f"{skor} / {total}")
                st.metric(label="Bu Testteki Başarı Oranı", value=f"%{int((skor / total) * 100)}")

                if st.button("Yeni Test Başlat"):
                    st.session_state.test_basladi = False
                    st.session_state.test_bitti = False
                    st.rerun()
