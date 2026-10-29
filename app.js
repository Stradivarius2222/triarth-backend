require('dotenv').config();
const express = require('express');
const path = require('path');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const fs = require('fs');
const { Resend } = require('resend');
const multer = require('multer');
const crypto = require('crypto');
const XLSX = require('xlsx');

// Import Client Turso (libSQL) menggantikan sql.js
const { createClient } = require('@libsql/client');

// ==========================================
// KONFIGURASI MULTER & UPLOAD FILE
// ==========================================
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, path.join(__dirname, 'public', 'uploads'));
    },
    filename: function (req, file, cb) {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        const ext = path.extname(file.originalname);
        cb(null, file.fieldname + '-' + uniqueSuffix + ext);
    }
});
const upload = multer({ storage: storage });

// ==========================================
// KONFIGURASI APLIKASI & RESEND
// ==========================================
const app = express();
const PORT = process.env.PORT || 3000;
const resend = new Resend(process.env.RESEND_API_KEY);

// ==========================================
// INISIALISASI DATABASE TURSO (LIBSQL)
// ==========================================
const db = createClient({
    url: process.env.TURSO_DATABASE_URL,
    authToken: process.env.TURSO_AUTH_TOKEN,
});

async function initDatabase() {
    try {
        await db.execute(`CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            company TEXT,
            email TEXT UNIQUE NOT NULL,
            phone TEXT,
            password TEXT NOT NULL,
            role TEXT DEFAULT 'customer',
            status TEXT DEFAULT 'active',
            reset_token TEXT,
            reset_expires DATETIME
        )`);

        await db.execute(`CREATE TABLE IF NOT EXISTS product_formulas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_name TEXT NOT NULL,
            segment TEXT NOT NULL,
            pdf_url TEXT NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        await db.execute(`CREATE TABLE IF NOT EXISTS news (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            slug TEXT UNIQUE NOT NULL,
            content TEXT NOT NULL,
            image_url TEXT,
            author_id INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (author_id) REFERENCES users(id)
        )`);

        await db.execute(`CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            segment TEXT NOT NULL,
            principal TEXT,
            product_name TEXT NOT NULL,
            slug TEXT UNIQUE NOT NULL,
            inci_name TEXT,
            description TEXT,
            function TEXT,
            application TEXT,
            formulation_detail TEXT,
            formula_pdf TEXT
        )`);

        await db.execute(`CREATE TABLE IF NOT EXISTS site_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )`);

        await db.execute(`CREATE TABLE IF NOT EXISTS partners (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            logo_url TEXT NOT NULL,
            website_url TEXT
        )`);

        await db.execute(`CREATE TABLE IF NOT EXISTS translations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            translation_key TEXT UNIQUE NOT NULL,
            en_text TEXT NOT NULL,
            id_text TEXT NOT NULL
        )`);

        await db.execute(`CREATE TABLE IF NOT EXISTS segments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            slug TEXT UNIQUE NOT NULL,
            description TEXT,
            banner_image TEXT
        )`);

        // Buat Akun Admin Default jika Belum Ada
        const hashedPass = await bcrypt.hash('admin123', 10);
        await db.execute({
            sql: `INSERT OR IGNORE INTO users (id, name, email, password, role, status) VALUES (1, 'Admin', 'admin@triarth.com', ?, 'admin', 'active')`,
            args: [hashedPass]
        });

        // Insert Default Translations
        const defaultTranslations = [
            ['aboutUs', 'About Us', 'Tentang Kami'],
            ['segments', 'Segments', 'Segmen'],
            ['segPersonal', 'Personal Care, Cosmetic & Fragrance', 'Perawatan Pribadi, Kosmetik & Wewangian'],
            ['segPharma', 'Pharmaceutical & Nutraceutical', 'Farmasi & Nutrasetikal'],
            ['segHome', 'Home Care I&I', 'Perawatan Rumah I&I'],
            ['segVet', 'Feedmill & Animal Health', 'Pakan Ternak & Kesehatan Hewan'],
            ['segProcess', 'Processing Chemicals', 'Bahan Kimia Proses'],
            ['news', 'News', 'Berita'],
            ['contact', 'Contact Us', 'Hubungi Kami'],
            ['signUp', 'Sign Up', 'Daftar'],
            ['signIn', 'Sign In', 'Masuk'],
            ['productFinderTitle', 'Product Finder', 'Pencari Produk'],
            ['searchPlaceholder', 'Search Product Name or INCI Name', 'Cari Nama Produk atau Nama INCI'],
            ['filterPrincipal', 'Principal', 'Prinsipal'],
            ['allPrincipals', 'All Principals', 'Semua Prinsipal'],
            ['filterFunction', 'Function', 'Fungsi'],
            ['allFunctions', 'All Functions', 'Semua Fungsi'],
            ['filterApplication', 'Application', 'Aplikasi'],
            ['allApplications', 'All Applications', 'Semua Aplikasi'],
            ['resetBtn', 'Reset', 'Atur Ulang'],
            ['thPrincipal', 'Principal', 'Prinsipal'],
            ['thProductName', 'Product Name', 'Nama Produk'],
            ['thInci', 'INCI Name', 'Nama INCI'],
            ['noDataMsg', 'No products found.', 'Tidak ada produk yang ditemukan.'],
            ['addressLabel', 'Address:', 'Alamat:'],
            ['navColTitle', 'Navigation', 'Navigasi'],
            ['workingHours', 'We are open: Mon - Fri: 8am - 5pm', 'Jam Kerja: Sen - Jum: 08:00 - 17:00'],
            ['rightsText', 'All rights reserved', 'Hak cipta dilindungi']
        ];

        for (const [key, enText, idText] of defaultTranslations) {
            await db.execute({
                sql: `INSERT OR IGNORE INTO translations (translation_key, en_text, id_text) VALUES (?, ?, ?)`,
                args: [key, enText, idText]
            });
        }

        // Insert Default Partners
        const defaultPartners = [
            ['Linqtec', '/images/partners/linqtec.png', 'https://www.linqtec.com/en/'],
            ['Corbion', '/images/partners/corbion_logo.png', 'https://www.corbion.com/biochemicals'],
            ['Evyap', '/images/partners/evyap.png', 'https://evyaplifechemistry.com/'],
            ['Kahlwax', '/images/partners/kahl.png', 'https://www.kahlworld.com/company/'],
            ['Sun Chemical', '/images/partners/sunchemical.jpg', 'https://www.sunchemical.com/'],
            ['Toyo Sugar', '/images/partners/toyosugar.jpg', 'https://www.toyosugar.co.jp/en/index.html'],
            ['Zagro', '/images/partners/zagro.jpg', 'https://www.zagro.com/'],
            ['BCF', '/images/partners/bcf.png', 'https://www.bcf-lifesciences.com/en/families-of-products/pharmaceutical/'],
            ['Trimak', '/images/partners/trimak_logo.png', 'https://www.triarth.com/'],
            ['Nucera', '/images/partners/nucera.png', 'https://nucerasolutions.com/'],
            ['DSM', '/images/partners/dsm.png', 'https://www.dsm-firmenich.com/en/home.html'],
            ['Lubrizol Beauty', '/images/partners/lubrizol_beauty.png', 'https://www.lubrizol.com/industries/beauty-and-personal-care'],
            ['Lubrizol Home', '/images/partners/lubrizol_home_logo_.png', 'https://www.lubrizol.com/industries/consumer-goods-and-home-care'],
            ['Aiglon', '/images/partners/aiglon.jpg', 'https://www.aiglon.eu/en/']
        ];

        for (const [name, logo, website] of defaultPartners) {
            await db.execute({
                sql: `INSERT OR IGNORE INTO partners (name, logo_url, website_url) VALUES (?, ?, ?)`,
                args: [name, logo, website]
            });
        }

        // Insert Default Settings
        const defaultSettings = [
            ['logo_white', '/images/logo2.png'],
            ['logo_dark', '/images/logo.png'],
            ['hero_title', 'The leading distributor of specialty chemicals in Indonesia'],
            ['who_we_are_title', 'Who We Are'],
            ['who_we_are_text', 'PT Tritunggal Arthamakmur is a premier distributor for top global manufacturers of specialty chemicals in Indonesia. We strive to provide an extensive choice of chemicals that meets international quality standards.'],
            ['cta_title', 'Professionalism means to be the best in the business.'],
            ['cta_text', 'Customer satisfaction is the goal. We provide exceptional customer service and are always ready to support you.'],
            ['footer_company_name', 'PT Tritunggal Arthamakmur'],
            ['footer_address', 'Komplek Graha Kencana Blok AB\nJl. Raya Pejuangan No. 88\nKebon Jeruk, Jakarta Barat\nJakarta 11530 Indonesia'],
            ['footer_phone', '+62 21 53677181 (Hunting)'],
            ['footer_working_hours', 'We are open: Mon - Fri: 8am - 5pm'],
            ['footer_linkedin', 'https://www.linkedin.com/company/tritunggalarthamakmur/home/'],
            ['footer_facebook', 'https://www.facebook.com/tritunggalarthamakmur'],
            ['footer_instagram', 'https://www.instagram.com/tritunggalarthamakmur/'],
            ['footer_copyright', '© 2026 Tritunggal Arthamakmur. All rights reserved'],
            ['news_newsletter_title', 'Get our free monthly newsletter.'],
            ['news_newsletter_text', 'Subscribe to get new products information and ideas for your next trendy products.'],
            ['contact_email_destination', 'rendy.riady24@gmail.com'],
            ['segment_processing_intro', 'We offer a wide range of specialty chemicals designed to optimize industrial processes, enhance product quality, and improve operational efficiency. Our portfolio includes functional chemical solutions tailored to meet the diverse needs of various industries, supporting applications such as process optimization, performance enhancement, and product quality improvement. With our commitment to quality and reliable supply, we work closely with our partners and customers to identify the right solutions for their specific processing needs, helping them achieve consistent results and greater operational efficiency.'],
            ['contact_heading', 'Let us know what you need'],
            ['principals_bg_image', '/images/principals-bg.jpg'],
            ['contact_subtext', 'Lets have a discussion. Reach out to us for any inquiries, partnerships, or more information. We are here to help you.'],
            ['principals_intro_text', 'We partner with reputable global manufacturers to bring quality ingredients, innovative technologies, and reliable solutions to the industries we serve. Trusted Partnerships. Proven Quality.'],
            ['about_hero_text', 'Building a healthier future through science, nature, and innovation. Discover our journey, our values, and our commitment to providing high-quality pharmaceutical and functional food solutions globality.'],
            ['news_hero_text', 'Explore the latest news, expert insights, product knowledge, practical tips, and company updates to keep you informed and inspired. Insights, Innovation & Industry Updates.'],
            ['hero_bg_image_1', '/images/slider banner-01.webp'],
            ['hero_bg_image_2', '/images/slider banner-02.webp'],
            ['hero_bg_image_3', '/images/slider banner-03.webp'],
            ['hero_bg_image_4', '/images/slider banner-04.webp'],
            ['hero_bg_image_5', '/images/slider banner-05.webp'],
            ['cta_bg_image', '/images/cta-bg.png'],
            ['contact_bg_image', '/images/contact-bg.webp'],
            ['news_bg_image', '/images/news-bg.webp'],
            ['newsletter_bg_image', '/images/news2.png'],
            ['about_bg_image', '/images/about-bg.jpg'],
            ['about_tab_img_1', '/images/about-us.png'],
            ['about_tab_img_2', '/images/history.png'],
            ['about_tab_img_3', '/images/vision-mission.png'],
            ['about_tab_img_4', '/images/presence.png'],
            ['segment_personal_care_bg', '/images/segments/personal-care-bg.jpg'],
            ['segment_pharma_bg', '/images/segments/pharma-bg.jpg'],
            ['segment_homecare_bg', '/images/segments/homecare-bg.jpg'],
            ['segment_vet_bg', '/images/segments/veterinary-bg.jpg'],
            ['segment_processing_bg', '/images/segments/processing-bg.jpg'],
            ['segment_personal_care_title', 'Personal Care, Cosmetic & Fragrance'],
            ['segment_personal_care_intro', 'We provide a comprehensive range of high-quality ingredients for personal care, cosmetics, and fragrance applications, covering hair care, skin care, color cosmetics, and fragrance formulations. Our portfolio includes rheology modifiers, emulsifiers, emollients, conditioning polymers, active ingredients, pearlescent pigments, preservatives, fragrance ingredients, and more. Backed by our experienced team and technical expertise, we work closely with our partners and customers to develop innovative formulations, enhance product performance, and bring new ideas to life.'],
            ['segment_pharma_title', 'Pharmaceutical & Nutraceutical'],
            ['segment_pharma_intro', 'We offer a comprehensive range of high-quality ingredients for pharmaceutical and nutraceutical applications, including Active Pharmaceutical Ingredients (APIs), pharmaceutical excipients, botanical extracts, vitamins, minerals, and other functional ingredients to support the development of pharmaceutical products, dietary supplements, and health-focused formulations. With our extensive product portfolio and commitment to quality, we work closely with our partners and customers to provide reliable ingredient solutions that support product innovation, formulation development, and evolving market needs.'],
            ['segment_homecare_title', 'Home Care I&I'],
            ['segment_homecare_intro', 'Supplying various specialty surfactants and performance additives such as hydrotrope, fabric softener, product thickener, emulsifier, dispersant, wetting agent, biocide and preservative to meet consumer needs for optimizing product performance for household cleaning and industrial & institutional applications.'],
            ['segment_vet_title', 'Feedmill & Animal Health'],
            ['segment_vet_intro', 'We provide a comprehensive range of specialty feed additives and animal health solutions designed to support animal nutrition, health, and performance. Our portfolio includes antioxidants, acidifiers, vitamins and mineral premixes, mycotoxin binders, mold inhibitors, and other functional ingredients for the feedmill, livestock, poultry, and aquaculture industries. Through our commitment to quality and reliable ingredient solutions, we work closely with our partners and customers to address industry challenges, optimize feed quality, support animal health, and improve overall production efficiency.'],
            ['segment_processing_title', 'Processing Chemicals'],
            ['processing_metalworking_desc', 'Offering specialty chemicals and customized blends to help Metalworking Fluid formulators create final products that are tailored to their specific needs with durable and high-performance solutions. We provide specialty additives for niche application areas in anti-foam, corrosion inhibitor, emulsifier, wetting agent, lubricant, biocide, and preservative.'],
            ['processing_battery_desc', 'Providing specialty additives for niche application areas in the battery manufacturing and processing.'],
            ['processing_textile_desc', 'Offering specialty chemicals for different steps of textile processing preparation, mercerization, bleaching, pre-treatment, finishing, washing off, and after-treatment, which includes specialty surfactants, levelling agents, and various softeners.'],
            ['processing_water_desc', 'Providing additives for high performance solutions in a broad case of applications for industrial water treatment.'],
            ['processing_cropcare_desc', 'Specialty surfactants, customized blends, and ingredients designed for Crop Care to help farmers prevent weeds, fungi, and pests, which destroy their crops.'],
            ['app_pc_cosmetics_title', 'Color Cosmetics'],
            ['app_pc_cosmetics_img', '/images/segments/cosmetic.png'],
            ['app_pc_skincare_title', 'Skin Care'],
            ['app_pc_skincare_img', '/images/segments/skincare.png'],
            ['app_pc_bath_title', 'Bath & Shower'],
            ['app_pc_bath_img', '/images/segments/bath.png'],
            ['app_pc_haircare_title', 'Hair Care'],
            ['app_pc_haircare_img', '/images/segments/haircare.png'],
            ['app_pharma_pharma_title', 'Pharmaceutical'],
            ['app_pharma_pharma_img', '/images/segments/pharma.png'],
            ['app_pharma_food_title', 'Functional Food'],
            ['app_pharma_food_img', '/images/segments/food.png'],
            ['app_pharma_excipient_title', 'Excipient'],
            ['app_pharma_excipient_img', '/images/segments/excipient.png'],
            ['app_homecare_home_title', 'Home Care'],
            ['app_homecare_home_img', '/images/segments/homecare1.png'],
            ['app_homecare_inst_title', 'Institutional Cleaning'],
            ['app_homecare_inst_img', '/images/segments/cleaning.png'],
            ['app_homecare_ind_title', 'Industrial Cleaning'],
            ['app_homecare_ind_img', '/images/segments/industrial.png'],
            ['app_vet_aqua_title', 'Aquaculture'],
            ['app_vet_aqua_img', '/images/segments/aqua.png'],
            ['app_vet_livestock_title', 'Livestock & Poultry'],
            ['app_vet_livestock_img', '/images/segments/livestock.png'],
            ['app_proc_metalworking_title', 'Metalworking Fluid'],
            ['app_proc_metalworking_img', '/images/segments/metalworking.png'],
            ['app_proc_battery_title', 'Battery'],
            ['app_proc_battery_img', '/images/segments/battery.png'],
            ['app_proc_textile_title', 'Textile Formulator'],
            ['app_proc_textile_img', '/images/segments/textile.png'],
            ['app_proc_water_title', 'Water Treatment'],
            ['app_proc_water_img', '/images/segments/water.png'],
            ['app_proc_cropcare_title', 'Crop Care Formulator'],
            ['app_proc_cropcare_img', '/images/segments/crop_care.png']
        ];

        for (const [key, val] of defaultSettings) {
            await db.execute({
                sql: `INSERT OR IGNORE INTO site_settings (key, value) VALUES (?, ?)`,
                args: [key, val]
            });
        }

        console.log('Inisialisasi Database Turso Berhasil.');
    } catch (err) {
        console.error('Gagal inisialisasi database Turso:', err.message);
    }
}

initDatabase();

// ==========================================
// MIDDLEWARE CONFIGURATION
// ==========================================
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
    secret: process.env.SESSION_SECRET || 'triarth-secret-key-default',
    resave: false,
    saveUninitialized: false,
    cookie: { 
        maxAge: 1000 * 60 * 60 * 8 
    }
}));

function isEditorOrAdmin(req) {
    return req.session.user && (req.session.user.role === 'editor' || req.session.user.role === 'admin');
}

function renderMessagePage(title, message, type = 'success', redirectUrl = '/', buttonText = 'Kembali') {
    const iconClass = type === 'success' ? 'fa-circle-check' : (type === 'warning' ? 'fa-triangle-exclamation' : 'fa-circle-xmark');
    const iconColor = type === 'success' ? '#008744' : (type === 'warning' ? '#d69e2e' : '#e53e3e');

    return `
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>${title} - PT Tritunggal Arthamakmur</title>
            <link rel="stylesheet" href="/css/style.css">
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                body { background-color: #f4f6f9; font-family: 'Open Sans', sans-serif; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }
                .message-card { background: #ffffff; max-width: 480px; width: 100%; padding: 45px 35px; border-radius: 16px; box-shadow: 0 10px 30px rgba(0, 0, 0, 0.08); text-align: center; box-sizing: border-box; }
                .message-icon { font-size: 3.5rem; color: ${iconColor}; margin-bottom: 20px; }
                .message-card h1 { font-size: 1.8rem; font-weight: 700; color: #1a1a1a; margin-bottom: 12px; }
                .message-card p { font-size: 1.05rem; color: #555555; line-height: 1.6; margin-bottom: 30px; }
                .btn-action { display: inline-block; background-color: #008744; color: #ffffff; text-decoration: none; padding: 12px 36px; border-radius: 30px; font-weight: 700; font-size: 1rem; }
            </style>
        </head>
        <body>
            <div class="message-card">
                <div class="message-icon"><i class="fa-solid ${iconClass}"></i></div>
                <h1>${title}</h1>
                <p>${message}</p>
                <a href="${redirectUrl}" class="btn-action">${buttonText}</a>
            </div>
        </body>
        </html>
    `;
}

// ==========================================
// ROUTE HALAMAN PUBLIK & STATIS
// ==========================================
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/principals', (req, res) => res.sendFile(path.join(__dirname, 'public', 'principals.html')));
app.get('/about', (req, res) => res.sendFile(path.join(__dirname, 'public', 'about.html')));
app.get('/news', (req, res) => res.sendFile(path.join(__dirname, 'public', 'news.html')));
app.get('/news-detail.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'news-detail.html')));
app.get('/contact', (req, res) => res.sendFile(path.join(__dirname, 'public', 'contact.html')));
app.get('/signup', (req, res) => res.sendFile(path.join(__dirname, 'public', 'signup.html')));
app.get('/signin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'signin.html')));
app.get('/segments/personal-care', (req, res) => res.sendFile(path.join(__dirname, 'public', 'segments', 'personal-care.html')));
app.get('/segments/pharmaceutical-functional-food', (req, res) => res.sendFile(path.join(__dirname, 'public', 'segments', 'pharmaceutical-functional-food.html')));
app.get('/segments/home-care-ii', (req, res) => res.sendFile(path.join(__dirname, 'public', 'segments', 'home-care-ii.html')));
app.get('/segments/veterinary-feedmill', (req, res) => res.sendFile(path.join(__dirname, 'public', 'segments', 'veterinary-feedmill.html')));
app.get('/segments/processing-chemicals', (req, res) => res.sendFile(path.join(__dirname, 'public', 'segments', 'processing-chemicals.html')));

// Contact Us
app.post('/contact', async (req, res) => {
    const { subject, message, email, 'g-recaptcha-response': recaptchaToken } = req.body;

    try {
        if (!recaptchaToken) {
            return res.status(400).send(renderMessagePage('Verifikasi Gagal', 'Silakan centang kotak reCAPTCHA terlebih dahulu.', 'warning', '/contact', 'Kembali'));
        }

        const verifyUrl = `https://www.google.com/recaptcha/api/siteverify?secret=${process.env.RECAPTCHA_SECRET_KEY}&response=${recaptchaToken}`;
        const recaptchaResponse = await fetch(verifyUrl, { method: 'POST' });
        const recaptchaResult = await recaptchaResponse.json();

        if (!recaptchaResult.success) {
            return res.status(400).send(renderMessagePage('Verifikasi Gagal', 'Validasi reCAPTCHA tidak valid. Silakan coba lagi.', 'error', '/contact', 'Kembali'));
        }

        let targetEmail = 'rendy.riady24@gmail.com'; 
        const settingRes = await db.execute(`SELECT value FROM site_settings WHERE key = 'contact_email_destination'`);
        if (settingRes.rows.length > 0 && settingRes.rows[0].value) {
            targetEmail = settingRes.rows[0].value;
        }

        await resend.emails.send({
            from: 'PT Tritunggal Arthamakmur <onboarding@resend.dev>',
            to: [targetEmail],
            subject: `Pesan Contact Us: ${subject}`,
            html: `
                <h3>Pesan Baru dari Website Contact Us</h3>
                <p><strong>Email Pengirim:</strong> ${email}</p>
                <p><strong>Subject:</strong> ${subject}</p>
                <p><strong>Pesan:</strong></p>
                <p>${message}</p>
            `
        });

        res.send(renderMessagePage('Pesan Berhasil Dikirim', 'Terima kasih telah menghubungi kami. Pesan Anda telah berhasil dikirim dan akan segera kami proses.', 'success', '/contact', 'Kembali'));

    } catch (err) {
        res.status(500).send(renderMessagePage('Error Server', err.message, 'error', '/contact', 'Kembali'));
    }
});

// Sign Up
app.post('/signup', async (req, res) => {
    const { name, company, email, phone, password, 
        'g-recaptcha-response': recaptchaToken 
    } = req.body;

    try {
        if (!recaptchaToken) {
            return res.status(400).send(renderMessagePage('Verifikasi Gagal', 'Silakan centang kotak reCAPTCHA terlebih dahulu.', 'warning', '/signup', 'Kembali'));
        }

        const verifyUrl = `https://www.google.com/recaptcha/api/siteverify?secret=${process.env.RECAPTCHA_SECRET_KEY}&response=${recaptchaToken}`;
        const recaptchaResponse = await fetch(verifyUrl, { method: 'POST' });
        const recaptchaResult = await recaptchaResponse.json();

        if (!recaptchaResult.success) {
            return res.status(400).send(renderMessagePage('Verifikasi Gagal', 'Validasi reCAPTCHA tidak valid. Silakan coba lagi.', 'error', '/signup', 'Kembali'));
        }

        const checkRes = await db.execute({
            sql: `SELECT COUNT(*) as count FROM users WHERE email = ?`,
            args: [email]
        });

        if (checkRes.rows[0].count > 0) {
            return res.status(400).send(renderMessagePage(
                'Email Sudah Terdaftar', 
                'Alamat email tersebut sudah terdaftar di sistem kami. Silakan gunakan email lain atau langsung masuk (Sign In).', 
                'warning', 
                '/signup', 
                'Coba Lagi'
            ));
        }

        const hashedPassword = await bcrypt.hash(password, 10);
        
        await db.execute({
            sql: `INSERT INTO users (name, company, email, phone, password, role, status) VALUES (?, ?, ?, ?, ?, ?, ?)`,
            args: [name, company, email, phone, hashedPassword, 'customer', 'pending']
        });

        res.send(renderMessagePage('Pendaftaran Berhasil', 'Akun Anda berhasil dibuat dan saat ini sedang menunggu persetujuan dari Administrator.', 'success', '/signin', 'Masuk (Sign In)'));

    } catch (err) {
        res.status(500).send(renderMessagePage('Error Server', err.message, 'error', '/signup', 'Kembali'));
    }
});

// Sign In
app.post('/signin', async (req, res) => {
    const { email, password, redirect, 'g-recaptcha-response': recaptchaToken } = req.body;
    
    try {
        if (!recaptchaToken) {
            return res.status(400).send(renderMessagePage('Verifikasi Gagal', 'Silakan centang kotak reCAPTCHA terlebih dahulu.', 'error', '/signin', 'Coba Lagi'));
        }

        const verifyUrl = `https://www.google.com/recaptcha/api/siteverify?secret=${process.env.RECAPTCHA_SECRET_KEY}&response=${recaptchaToken}`;
        const recaptchaResponse = await fetch(verifyUrl, { method: 'POST' });
        const recaptchaResult = await recaptchaResponse.json();

        if (!recaptchaResult.success) {
            return res.status(400).send(renderMessagePage('Verifikasi Gagal', 'Validasi reCAPTCHA tidak valid. Silakan coba lagi.', 'error', '/signin', 'Coba Lagi'));
        }

        const userRes = await db.execute({
            sql: `SELECT * FROM users WHERE email = ?`,
            args: [email]
        });

        const user = userRes.rows[0];

        if (!user || !(await bcrypt.compare(password, user.password))) {
            return res.status(400).send(renderMessagePage('Login Gagal', 'Email atau password salah.', 'error', '/signin', 'Coba Lagi'));
        }
        if (user.role === 'editor' && user.status === 'pending') {
            return res.status(403).send(renderMessagePage('Pending', 'Akun Editor masih menunggu persetujuan Admin.', 'warning', '/signin', 'Kembali'));
        }

        req.session.user = { id: user.id, name: user.name, role: user.role };
        
        if (user.role === 'admin') {
            res.redirect('/admin/dashboard');
        } else if (user.role === 'editor') {
            res.redirect('/editor/dashboard');
        } else {
            const destination = (redirect && redirect.startsWith('/')) ? redirect : '/';
            res.redirect(destination);
        }

    } catch (err) {
        res.status(500).send(renderMessagePage('Error', 'Kesalahan server saat login: ' + err.message, 'error', '/signin', 'Kembali'));
    }
});

// ==========================================
// FORGOT PASSWORD & RESET PASSWORD
// ==========================================
app.get('/forgot-password', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'forgot-password.html'));
});

app.post('/forgot-password', async (req, res) => {
    const { email, 'g-recaptcha-response': recaptchaToken } = req.body;

    try {
        if (!recaptchaToken) {
            return res.status(400).send(renderMessagePage('Verifikasi Gagal', 'Silakan centang reCAPTCHA terlebih dahulu.', 'warning', '/forgot-password', 'Kembali'));
        }

        const verifyUrl = `https://www.google.com/recaptcha/api/siteverify?secret=${process.env.RECAPTCHA_SECRET_KEY}&response=${recaptchaToken}`;
        const recaptchaResponse = await fetch(verifyUrl, { method: 'POST' });
        const recaptchaResult = await recaptchaResponse.json();

        if (!recaptchaResult.success) {
            return res.status(400).send(renderMessagePage('Verifikasi Gagal', 'Validasi reCAPTCHA tidak valid.', 'error', '/forgot-password', 'Kembali'));
        }

        const userRes = await db.execute({
            sql: `SELECT * FROM users WHERE email = ?`,
            args: [email]
        });
        const user = userRes.rows[0];

        if (!user) {
            return res.send(renderMessagePage('Email Terkirim', 'Jika email terdaftar, instruksi reset password telah dikirim ke email Anda.', 'success', '/signin', 'Kembali ke Sign In'));
        }

        const resetToken = crypto.randomBytes(32).toString('hex');
        const resetExpires = new Date(Date.now() + 3600000).toISOString();

        await db.execute({
            sql: `UPDATE users SET reset_token = ?, reset_expires = ? WHERE id = ?`,
            args: [resetToken, resetExpires, user.id]
        });

        const resetLink = `http://${req.headers.host}/reset-password?token=${resetToken}`;
        
        await resend.emails.send({
            from: 'PT Tritunggal Arthamakmur <onboarding@resend.dev>',
            to: [email],
            subject: 'Reset Password - PT Tritunggal Arthamakmur',
            html: `
                <h3>Permintaan Reset Password</h3>
                <p>Halo <strong>${user.name}</strong>,</p>
                <p>Kami menerima permintaan untuk mereset password akun Anda. Klik tombol di bawah ini untuk membuat password baru:</p>
                <p><a href="${resetLink}" style="background-color:#008744; color:#fff; padding:10px 20px; text-decoration:none; border-radius:20px; font-weight:bold; display:inline-block;">Reset Password</a></p>
                <p>Link ini berlaku selama 1 jam.</p>
                <p>Jika Anda tidak meminta reset password, abaikan email ini.</p>
            `
        });

        res.send(renderMessagePage('Email Terkirim', 'Instruksi reset password telah dikirim ke email Anda. Silakan periksa kotak masuk/spam Anda.', 'success', '/signin', 'Kembali ke Sign In'));

    } catch (err) {
        res.status(500).send(renderMessagePage('Error Server', err.message, 'error', '/forgot-password', 'Kembali'));
    }
});

app.get('/reset-password', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'reset-password.html'));
});

app.post('/reset-password', async (req, res) => {
    const { token, new_password, confirm_password } = req.body;

    if (new_password !== confirm_password) {
        return res.status(400).send(renderMessagePage('Gagal Reset', 'Konfirmasi password tidak cocok.', 'error', `/reset-password?token=${token}`, 'Coba Lagi'));
    }

    try {
        const userRes = await db.execute({
            sql: `SELECT * FROM users WHERE reset_token = ? AND reset_expires > ?`,
            args: [token, new Date().toISOString()]
        });
        const user = userRes.rows[0];

        if (!user) {
            return res.status(400).send(renderMessagePage('Token Tidak Valid', 'Link reset password sudah kadaluarsa atau tidak valid.', 'error', '/forgot-password', 'Minta Link Baru'));
        }

        const hashedPassword = await bcrypt.hash(new_password, 10);
        await db.execute({
            sql: `UPDATE users SET password = ?, reset_token = NULL, reset_expires = NULL WHERE id = ?`,
            args: [hashedPassword, user.id]
        });

        res.send(renderMessagePage('Password Berhasil Diubah', 'Password Anda telah berhasil diperbarui. Silakan login dengan password baru Anda.', 'success', '/signin', 'Sign In'));

    } catch (err) {
        res.status(500).send(renderMessagePage('Error Server', err.message, 'error', '/forgot-password', 'Kembali'));
    }
});

// ==========================================
// EDITOR DASHBOARD & NEWS MANAGEMENT
// ==========================================
app.get('/editor/dashboard', async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.redirect('/signin');
    
    const result = await db.execute(`
        SELECT news.id, news.title, news.created_at, users.name as author_name 
        FROM news 
        JOIN users ON news.author_id = users.id 
        ORDER BY news.created_at DESC
    `);
    
    const allNews = result.rows;

    let newsRows = allNews.map(n => `
        <tr>
            <td style="padding: 12px; border: 1px solid #e2e8f0;">${n.title}</td>
            <td style="padding: 12px; border: 1px solid #e2e8f0;">${n.author_name}</td>
            <td style="padding: 12px; border: 1px solid #e2e8f0;">${n.created_at}</td>
            <td style="padding: 12px; border: 1px solid #e2e8f0; text-align: center;">
                <a href="/editor/news/edit/${n.id}" style="background:#3182ce; color:#fff; padding:6px 12px; text-decoration:none; border-radius:4px; font-weight:600; margin-right:5px; display:inline-block;">Edit</a>
                <button type="button" onclick="openDeleteNewsModal(${n.id}, '${String(n.title).replace(/<[^>]*>?/gm, '').replace(/'/g, "\\'")}')" style="background:#e53e3e; color:#fff; border:none; padding:6px 14px; border-radius:6px; cursor:pointer; font-weight:600;">Delete</button>
            </td>
        </tr>
    `).join('');

    const topBackLink = req.session.user.role === 'admin' 
        ? '<a href="/admin/dashboard" style="background: #718096; color: #fff; padding: 8px 16px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 0.9rem; display: inline-flex; align-items: center; gap: 6px;"><i class="fa-solid fa-arrow-left"></i> Kembali ke Dashboard</a>' 
        : '';

    res.send(`
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Berita Dashboard - PT Tritunggal Arthamakmur</title>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
            <link href="https://cdn.jsdelivr.net/npm/summernote@0.8.18/dist/summernote-lite.min.css" rel="stylesheet">
            <script src="https://code.jquery.com/jquery-3.4.1.slim.min.js" crossorigin="anonymous"></script>
            <script src="https://cdn.jsdelivr.net/npm/summernote@0.8.18/dist/summernote-lite.min.js"></script>
            <style>
                body { font-family: 'Open Sans', sans-serif; background: #f4f6f9; margin: 0; padding: 40px; color: #333; }
                .container { max-width: 950px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
                .header-flex { display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; flex-wrap: wrap; gap: 15px; }
                h1 { color: #009639; margin: 0; }
                .form-group { margin-bottom: 20px; }
                label { display: block; font-weight: 600; margin-bottom: 8px; color: #2d3748; }
                input[type="text"], textarea { 
                    width: 100%; 
                    padding: 10px; 
                    border: 1px solid #cbd5e0; 
                    border-radius: 6px; 
                    box-sizing: border-box; 
                    font-family: 'Open Sans', sans-serif; 
                    font-size: 0.95rem;
                }
                textarea { height: 130px; }
                button[type="submit"] { background: #009639; color: #fff; border: none; padding: 12px 24px; font-weight: 600; border-radius: 6px; cursor: pointer; }
                .table-responsive { width: 100%; overflow-x: auto; margin-top: 20px; }
                table { width: 100%; border-collapse: collapse; white-space: nowrap; }
                th, td { padding: 12px; border: 1px solid #e2e8f0; text-align: left; }
                th { background-color: #f8f9fa; color: #2d3748; }
                .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); justify-content: center; align-items: center; z-index: 1000; }
                .modal-overlay.active { display: flex; }
                .modal-box { background: #fff; width: 100%; max-width: 400px; padding: 30px; border-radius: 8px; text-align: center; box-sizing: border-box; }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header-flex">
                    <h1>Manajemen Berita</h1>
                    ${topBackLink}
                </div>
                <p style="margin-top: 5px; margin-bottom: 25px;">Selamat datang <strong>${req.session.user.name}</strong> (${req.session.user.role.toUpperCase()}).</p>
                
                <h3>Buat Berita / Artikel Baru</h3>
                <form action="/editor/news" method="POST" enctype="multipart/form-data">
                <div class="form-group">
                    <label>Judul Berita</label>
                    <textarea name="title" id="summernoteTitle" required placeholder="Masukkan judul berita..."></textarea>
                </div>
                    <div class="form-group">
                        <label>Unggah Gambar</label>
                        <input type="file" name="image" accept="image/*">
                    </div>
                    <div class="form-group">
                        <label>Konten Berita</label>
                        <textarea name="content" id="summernote" required></textarea>
                    </div>
                    <button type="submit">Publikasikan Berita</button>
                </form>

                <h3 style="margin-top: 45px;">Riwayat Berita</h3>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr><th>Judul Berita</th><th>Penulis</th><th>Tanggal Publikasi</th><th style="text-align: center; width: 160px;">Aksi</th></tr>
                        </thead>
                        <tbody>
                            ${newsRows.length > 0 ? newsRows : '<tr><td colspan="4" style="text-align:center; padding: 20px;">Belum ada berita.</td></tr>'}
                        </tbody>
                    </table>
                </div>
            </div>

            <div id="deleteNewsModal" class="modal-overlay">
                <div class="modal-box">
                    <h3>Konfirmasi Hapus Berita</h3>
                    <p id="deleteNewsModalText">Apakah Anda yakin ingin menghapus berita ini?</p>
                    <div style="display:flex; justify-content:center; gap:10px; margin-top:20px;">
                        <button type="button" onclick="closeDeleteNewsModal()" style="padding:8px 16px; border-radius:4px; border:none; cursor:pointer;">Batal</button>
                        <form id="deleteNewsForm" method="POST"><button type="submit" style="background:#e53e3e; color:#fff; padding:8px 16px; border-radius:4px; border:none; cursor:pointer;">Hapus</button></form>
                    </div>
                </div>
            </div>

            <script>
                function openDeleteNewsModal(id, title) {
                    document.getElementById('deleteNewsModalText').innerHTML = 'Hapus berita <strong>"' + title + '"</strong>?';
                    document.getElementById('deleteNewsForm').action = '/editor/news/delete/' + id;
                    document.getElementById('deleteNewsModal').classList.add('active');
                }
                function closeDeleteNewsModal() {
                    document.getElementById('deleteNewsModal').classList.remove('active');
                }
            </script>
            <script>
                $('#summernote').summernote({
                    placeholder: 'Tulis atau paste isi berita di sini...',
                    tabsize: 2,
                    height: 300
                });
            </script>
            <script>
                $('#summernoteTitle').summernote({
                    placeholder: 'Tulis atau paste judul berita di sini...',
                    tabsize: 2,
                    height: 100
                });
            </script>
        </body>
        </html>
    `);
});

app.post('/editor/news', upload.single('image'), async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    const { title, content } = req.body;
    const cleanTitleForSlug = title.replace(/<[^>]*>?/gm, '');
    const slug = cleanTitleForSlug.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') + '-' + Date.now();
    const imageUrl = req.file ? `/uploads/${req.file.filename}` : '';

    try {
        await db.execute({
            sql: `INSERT INTO news (title, slug, content, image_url, author_id) VALUES (?, ?, ?, ?, ?)`,
            args: [title, slug, content, imageUrl, req.session.user.id]
        });
        res.redirect('/editor/dashboard');
    } catch (err) {
        res.status(500).send('Gagal menyimpan berita: ' + err.message);
    }
});

app.get('/editor/news/edit/:id', async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.redirect('/signin');
    
    const result = await db.execute({
        sql: `SELECT * FROM news WHERE id = ?`,
        args: [req.params.id]
    });
    
    const article = result.rows[0];
    if (!article) return res.status(404).send('Berita tidak ditemukan.');

    res.send(`
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="UTF-8"><title>Edit Berita</title>
            <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
            <link href="https://cdn.jsdelivr.net/npm/summernote@0.8.18/dist/summernote-lite.min.css" rel="stylesheet">
            <script src="https://code.jquery.com/jquery-3.4.1.slim.min.js"></script>
            <script src="https://cdn.jsdelivr.net/npm/summernote@0.8.18/dist/summernote-lite.min.js"></script>
            <style>
                body { font-family: 'Open Sans', sans-serif; background: #f4f6f9; margin: 0; padding: 40px; color: #333; }
                .container { max-width: 800px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; }
                .form-group { margin-bottom: 20px; }
                label { display: block; font-weight: 600; margin-bottom: 8px; }
                input[type="text"], textarea { width: 100%; padding: 10px; border: 1px solid #cbd5e0; border-radius: 6px; }
                textarea { height: 160px; }
                button[type="submit"] { background: #3182ce; color: #fff; border: none; padding: 12px 24px; border-radius: 6px; cursor: pointer; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Edit Berita</h1>
                <form action="/editor/news/edit/${article.id}" method="POST" enctype="multipart/form-data">
                <div class="form-group">
                    <label>Judul Berita</label>
                    <textarea name="title" id="summernoteTitle" required>${article.title}</textarea>
                </div>
                    <div class="form-group">
                        <label>Ganti Gambar (Opsional)</label>
                        ${article.image_url ? `<div style="margin-bottom:10px;"><img src="${article.image_url}" style="max-width:150px; border-radius:6px;"></div>` : ''}
                        <input type="file" name="image" accept="image/*">
                    </div>
                    <div class="form-group">
                        <label>Konten Berita</label>
                        <textarea name="content" id="summernoteEdit" required>${article.content}</textarea>
                    </div>
                    <button type="submit">Simpan Perubahan</button>
                </form>
                <a href="/editor/dashboard" style="display:inline-block; margin-top:20px; color:#718096;">Kembali</a>
            </div>
        </body>
        <script>
            $(document).ready(function() {
                $('#summernoteEdit').summernote({ height: 250, placeholder: 'Tulis atau paste isi berita...' });
                $('#summernoteTitle').summernote({ height: 100, placeholder: 'Tulis atau paste judul berita...' });
            });
        </script>
        </html>
    `);
});

app.post('/editor/news/edit/:id', upload.single('image'), async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    const { title, content } = req.body;
    const newsId = req.params.id;

    try {
        if (req.file) {
            const imageUrl = `/uploads/${req.file.filename}`;
            await db.execute({
                sql: `UPDATE news SET title = ?, content = ?, image_url = ? WHERE id = ?`,
                args: [title, content, imageUrl, newsId]
            });
        } else {
            await db.execute({
                sql: `UPDATE news SET title = ?, content = ? WHERE id = ?`,
                args: [title, content, newsId]
            });
        }
        res.redirect('/editor/dashboard');
    } catch (err) {
        res.status(500).send('Gagal memperbarui berita: ' + err.message);
    }
});

app.post('/editor/news/delete/:id', async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    try {
        await db.execute({
            sql: `DELETE FROM news WHERE id = ?`,
            args: [req.params.id]
        });
        res.redirect('/editor/dashboard');
    } catch (err) {
        res.status(500).send('Gagal menghapus berita: ' + err.message);
    }
});

// ==========================================
// ADMIN DASHBOARD
// ==========================================
app.get('/admin/dashboard', async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/signin');
    
    const userRes = await db.execute(`SELECT id, name, email, role, status FROM users`);
    const users = userRes.rows;
    
    let userRows = users.map((u, index) => {
        const isSelf = req.session.user && req.session.user.id === u.id;
        const canSelect = !isSelf;

        return `
            <tr>
                <td style="text-align: center;">
                    ${canSelect ? `<input type="checkbox" name="selected_user_ids" value="${u.id}" form="userBulkForm" class="user-checkbox" onchange="updateUserBulkBtn()">` : ''}
                </td>
                <td>${index + 1}</td>
                <td>${u.name}</td>
                <td>${u.email}</td>
                <td>
                    ${u.role === 'admin' ? '<strong>Admin</strong>' : `<span style="text-transform: capitalize; font-weight: 600; color: #2b6cb0;">${u.role}</span>`}
                </td>
                <td><span style="color: ${u.status === 'pending' ? '#d69e2e' : '#008744'}; font-weight: bold;">${u.status}</span></td>
                <td>
                    ${u.status === 'pending' ? `
                        <form action="/admin/approve/${u.id}" method="POST" style="display:inline;">
                            <button type="submit" style="background:#008744; color:#fff; border:none; padding:6px 14px; border-radius:6px; cursor:pointer;">Approve</button>
                        </form>
                    ` : ''}
                    ${!isSelf && u.role !== 'admin' ? `
                        <button type="button" onclick="openEditRoleModal(${u.id}, '${String(u.name).replace(/'/g, "\\'")}', '${u.role}')" style="background:#3182ce; color:#fff; border:none; padding:6px 12px; border-radius:6px; cursor:pointer; font-weight:600; font-size:0.85rem; margin-left:5px;">
                            <i class="fa-solid fa-pen-to-square"></i> Edit Role
                        </button>
                    ` : ''}
                    ${!isSelf ? `
                        <button type="button" onclick="openDeleteUserModal(${u.id}, '${String(u.name).replace(/'/g, "\\'")}')" style="background:#e53e3e; color:#fff; border:none; padding:6px 14px; border-radius:6px; cursor:pointer; font-weight:600; font-size:0.85rem; margin-left:5px;">Delete</button>
                    ` : '<span style="color: #718096; font-size: 0.85rem; font-style: italic; margin-left:5px;">Akun Anda</span>'}
                </td>
            </tr>
        `;
    }).join('');

    res.send(`
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Admin Dashboard - PT Tritunggal Arthamakmur</title>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Open Sans', sans-serif; background: #f4f6f9; margin: 0; padding: 40px; color: #333; box-sizing: border-box; }
                .container { max-width: 1050px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
                h1 { color: #009639; margin-top: 0; }
                .section-box { background: #fff; padding: 25px; border-radius: 10px; margin-top: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
                .table-responsive { width: 100%; overflow-x: auto; margin-top: 15px; }
                table { width: 100%; border-collapse: collapse; white-space: nowrap; }
                th, td { padding: 12px; border: 1px solid #e2e8f0; text-align: left; font-size: 0.95rem; }
                th { background-color: #f8f9fa; color: #2d3748; }
                .logout-btn { display: inline-block; margin-top: 25px; background: #718096; color: #fff; padding: 8px 18px; text-decoration: none; border-radius: 6px; font-weight: 600; }

                .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.5); justify-content: center; align-items: center; z-index: 1000; }
                .modal-overlay.active { display: flex; }
                .modal-card { background: #fff; width: 100%; max-width: 420px; padding: 30px; border-radius: 12px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.15); box-sizing: border-box; }
                .modal-icon { font-size: 3rem; color: #e53e3e; margin-bottom: 15px; }
                .modal-card h3 { margin: 0 0 10px 0; font-size: 1.3rem; color: #1a202c; }
                .modal-card p { color: #4a5568; font-size: 0.95rem; margin-bottom: 25px; line-height: 1.5; }
                .modal-actions { display: flex; gap: 12px; justify-content: center; }
                .btn-modal-cancel { background: #e2e8f0; color: #2d3748; border: none; padding: 10px 20px; border-radius: 6px; font-weight: 600; cursor: pointer; flex: 1; }
                .btn-modal-confirm { background: #e53e3e; color: #fff; border: none; padding: 10px 20px; border-radius: 6px; font-weight: 600; cursor: pointer; flex: 1; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Admin Dashboard</h1>
                <p>Selamat datang Admin, <strong>${req.session.user.name}</strong>.</p>
                
                <div style="display: flex; gap: 12px; margin-bottom: 25px; flex-wrap: wrap;">
                    <a href="/editor/dashboard" style="background:#009639; color:#fff; padding:10px 18px; text-decoration:none; border-radius:6px; font-weight:600;">
                        <i class="fa-solid fa-newspaper" style="margin-right:6px;"></i> Kelola Berita & Artikel
                    </a>
                    <a href="/admin/products/dashboard" style="background:#2b6cb0; color:#fff; padding:10px 18px; text-decoration:none; border-radius:6px; font-weight:600;">
                        <i class="fa-solid fa-boxes-stacked" style="margin-right:6px;"></i> Kelola Produk & Upload
                    </a>
                    <a href="/admin/cms" style="background:#3182ce; color:#fff; padding:10px 18px; text-decoration:none; border-radius:6px; font-weight:600;">
                        <i class="fa-solid fa-pen-to-square" style="margin-right:6px;"></i> Kelola Konten Website (CMS)
                    </a>
                </div>

                <div class="section-box">
                    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 10px;">
                        <h3 style="color: #008744; margin: 0;"><i class="fa-solid fa-users"></i> Kelola User</h3>
                        <button type="button" id="userBulkDeleteBtn" onclick="openUserBulkModal()" style="background:#e53e3e; color:#fff; border:none; padding:7px 14px; border-radius:6px; cursor:pointer; font-weight:600; font-size: 0.9rem; display:none;">
                            <i class="fa-solid fa-trash"></i> Hapus Pilihan (<span id="userSelectedCount">0</span>)
                        </button>
                    </div>

                    <form id="userBulkForm" action="/admin/users/bulk-delete" method="POST"></form>

                    <div class="table-responsive">
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 40px; text-align: center;"><input type="checkbox" id="selectAllUsers" onclick="toggleSelectAllUsers(this)"></th>
                                    <th>ID</th>
                                    <th>Nama</th>
                                    <th>Email</th>
                                    <th>Role</th>
                                    <th>Status</th>
                                    <th>Aksi</th>
                                </tr>
                            </thead>
                            <tbody>${userRows}</tbody>
                        </table>
                    </div>
                </div>

                <a href="/logout" class="logout-btn">Logout</a>
            </div>

            <div id="customDeleteModal" class="modal-overlay">
                <div class="modal-card">
                    <div class="modal-icon"><i class="fa-solid fa-triangle-exclamation"></i></div>
                    <h3 id="modalTitle">Konfirmasi Hapus</h3>
                    <p id="modalDeleteMessage">Apakah Anda yakin ingin menghapus user ini?</p>
                    <div class="modal-actions">
                        <button type="button" class="btn-modal-cancel" onclick="closeCustomModal()">Batal</button>
                        <button type="button" class="btn-modal-confirm" id="modalConfirmBtn" onclick="executeDeleteAction()">Ya, Hapus</button>
                    </div>
                </div>
            </div>

            <form id="singleDeleteForm" method="POST" style="display: none;"></form>

            <script>
                let deleteType = '';
                let singleUserId = null;

                function toggleSelectAllUsers(source) {
                    const checkboxes = document.querySelectorAll('.user-checkbox');
                    checkboxes.forEach(cb => cb.checked = source.checked);
                    updateUserBulkBtn();
                }

                function updateUserBulkBtn() {
                    const checkedCount = document.querySelectorAll('.user-checkbox:checked').length;
                    const btn = document.getElementById('userBulkDeleteBtn');
                    const countSpan = document.getElementById('userSelectedCount');
                    
                    if (countSpan) countSpan.textContent = checkedCount;
                    if (btn) {
                        btn.style.display = checkedCount > 0 ? 'inline-block' : 'none';
                    }
                }

                function openDeleteUserModal(userId, userName) {
                    deleteType = 'single';
                    singleUserId = userId;
                    document.getElementById('modalTitle').textContent = 'Konfirmasi Hapus User';
                    document.getElementById('modalDeleteMessage').innerHTML = 'Apakah Anda yakin ingin menghapus akun <strong>"' + userName + '"</strong>?';
                    document.getElementById('customDeleteModal').classList.add('active');
                }

                function openUserBulkModal() {
                    deleteType = 'bulk';
                    const count = document.querySelectorAll('.user-checkbox:checked').length;
                    document.getElementById('modalTitle').textContent = 'Konfirmasi Hapus Massal';
                    document.getElementById('modalDeleteMessage').innerHTML = 'Apakah Anda yakin ingin menghapus <strong>' + count + ' user</strong> yang dipilih?';
                    document.getElementById('customDeleteModal').classList.add('active');
                }

                function closeCustomModal() {
                    document.getElementById('customDeleteModal').classList.remove('active');
                }

                function executeDeleteAction() {
                    if (deleteType === 'single') {
                        const form = document.getElementById('singleDeleteForm');
                        form.action = '/admin/delete/' + singleUserId;
                        form.submit();
                    } else if (deleteType === 'bulk') {
                        document.getElementById('userBulkForm').submit();
                    }
                }
                function openEditRoleModal(userId, userName, currentRole) {
                    document.getElementById('editRoleForm').action = '/admin/update-role/' + userId;
                    document.getElementById('editRoleUserName').innerHTML = 'Mengubah role untuk akun: <strong>"' + userName + '"</strong>';
                    document.getElementById('userRoleSelect').value = currentRole;
                    document.getElementById('editRoleModal').classList.add('active');
                }

                function closeEditRoleModal() {
                    document.getElementById('editRoleModal').classList.remove('active');
                }
            </script>
            
            <div id="editRoleModal" class="modal-overlay">
                <div class="modal-card" style="max-width: 400px; text-align: left;">
                    <h3 style="margin-top: 0; color: #008744; display: flex; align-items: center; gap: 8px;">
                        <i class="fa-solid fa-pen-to-square"></i> Ubah Role User
                    </h3>
                    <p id="editRoleUserName" style="color: #4a5568; font-size: 0.95rem; margin-bottom: 15px;"></p>
                    <form id="editRoleForm" method="POST">
                        <div style="margin-bottom: 20px;">
                            <label style="display:block; font-weight:600; margin-bottom:8px; font-size:0.9rem;">Pilih Role:</label>
                            <select id="userRoleSelect" name="role" required style="width: 100%; padding: 10px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box; font-size: 0.95rem;">
                                <option value="customer">Customer</option>
                                <option value="editor">Editor</option>
                            </select>
                        </div>
                        <div style="display: flex; gap: 10px; justify-content: flex-end;">
                            <button type="button" class="btn-modal-cancel" onclick="closeEditRoleModal()" style="flex: none; padding: 8px 16px;">Batal</button>
                            <button type="submit" style="background: #008744; color: #fff; border: none; padding: 8px 20px; border-radius: 6px; font-weight: 600; cursor: pointer;">Simpan Perubahan</button>
                        </div>
                    </form>
                </div>
            </div>
        </body>
        </html>
    `);
});

app.post('/admin/users/bulk-delete', async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send('Unauthorized');
    try {
        let { selected_user_ids } = req.body;
        if (!selected_user_ids) {
            return res.status(400).send('Tidak ada user yang dipilih.');
        }

        if (!Array.isArray(selected_user_ids)) {
            selected_user_ids = [selected_user_ids];
        }

        selected_user_ids = selected_user_ids.filter(id => parseInt(id) !== req.session.user.id);

        if (selected_user_ids.length > 0) {
            const placeholders = selected_user_ids.map(() => '?').join(',');
            await db.execute({
                sql: `DELETE FROM users WHERE id IN (${placeholders})`,
                args: selected_user_ids
            });
        }

        res.redirect('/admin/dashboard');
    } catch (err) {
        res.status(500).send('Gagal menghapus user secara massal: ' + err.message);
    }
});

app.post('/admin/update-role/:id', async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/signin');
    const { role } = req.body;
    try {
        await db.execute({
            sql: `UPDATE users SET role = ? WHERE id = ?`,
            args: [role, req.params.id]
        });
        res.redirect('/admin/dashboard');
    } catch (err) {
        res.status(500).send('Gagal memperbarui role: ' + err.message);
    }
});

app.get('/admin/update-role/:id', (req, res) => res.redirect('/admin/dashboard'));

app.post('/admin/approve/:id', async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send('Unauthorized');
    try {
        await db.execute({
            sql: `UPDATE users SET status = 'active' WHERE id = ?`,
            args: [req.params.id]
        });
        res.redirect('/admin/dashboard');
    } catch (err) {
        res.status(500).send('Gagal menyetujui akun: ' + err.message);
    }
});

app.post('/admin/delete/:id', async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send('Unauthorized');
    if (req.session.user.id === parseInt(req.params.id)) {
        return res.status(400).send('Anda tidak dapat menghapus akun Anda sendiri.');
    }
    try {
        await db.execute({
            sql: `DELETE FROM users WHERE id = ?`,
            args: [req.params.id]
        });
        res.redirect('/admin/dashboard');
    } catch (err) {
        res.status(500).send('Gagal menghapus akun: ' + err.message);
    }
});

// ==========================================
// KELOLA PRODUK DASHBOARD
// ==========================================
app.get('/admin/products/dashboard', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/signin');

    const segments = [
        "Personal Care, Cosmetic & Fragrance",
        "Pharmaceutical & Nutraceutical",
        "Home Care I&I",
        "Feedmill & Animal Health",
        "Processing Chemicals"
    ];

    res.send(`
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Kelola Produk - PT Tritunggal Arthamakmur</title>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Open Sans', sans-serif; background: #f4f6f9; margin: 0; padding: 40px; color: #333; box-sizing: border-box; }
                .container { max-width: 900px; margin: 0 auto; background: #fff; padding: 35px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
                h1 { color: #009639; margin-top: 0; font-size: 1.8rem; display: flex; align-items: center; gap: 10px; }
                .section-box { background: #fff; padding: 25px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; }
                .btn-back { display: inline-block; background: #718096; color: #fff; padding: 9px 18px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 0.9rem; margin-bottom: 20px; }
                .btn-back:hover { background: #4a5568; }
                .segment-btn { display: flex; align-items: center; justify-content: space-between; background: #f8f9fa; border: 1px solid #cbd5e0; padding: 15px 20px; border-radius: 8px; text-decoration: none; color: #2d3748; font-weight: 600; font-size: 1.05rem; transition: all 0.2s; margin-bottom: 12px; }
                .segment-btn:hover { background: #008744; color: #fff; border-color: #008744; transform: translateY(-2px); box-shadow: 0 4px 10px rgba(0,0,0,0.08); }
                .segment-btn:hover i { color: #fff !important; }
            </style>
        </head>
        <body>
            <div class="container">
                <a href="/admin/dashboard" class="btn-back"><i class="fa-solid fa-arrow-left"></i> Kembali ke Admin Dashboard</a>
                <h1><i class="fa-solid fa-boxes-stacked"></i> Kelola Produk & Unggahan</h1>
                <p>Silakan lakukan unggah data atau pilih segmen di bawah untuk mengelola produk dan riwayat formula.</p>

                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 25px;">
                    <div class="section-box" style="margin-top: 0;">
                        <h3 style="color: #008744; margin-top: 0;"><i class="fa-solid fa-file-excel"></i> Upload Excel / CSV</h3>
                        <form action="/admin/products/upload-by-segment" method="POST" enctype="multipart/form-data" style="margin-top: 15px;">
                            <div style="margin-bottom: 12px;">
                                <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Segmen Default:</label>
                                <select name="segment" required style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                                    ${segments.map(s => `<option value="${s}">${s}</option>`).join('')}
                                </select>
                            </div>
                            <div style="margin-bottom: 12px;">
                                <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">File Excel:</label>
                                <input type="file" name="product_file" accept=".xlsx, .xls, .csv" required style="padding: 6px; background: #f8f9fa; border: 1px dashed #cbd5e0; border-radius: 6px; width: 100%; box-sizing: border-box;">
                            </div>
                            <button type="submit" style="background: #008744; color: #fff; border: none; padding: 9px 16px; font-weight: 600; border-radius: 6px; cursor: pointer; font-size: 0.9rem;">Upload Produk</button>
                        </form>
                    </div>

                    <div class="section-box" style="margin-top: 0;">
                        <h3 style="color: #008744; margin-top: 0;"><i class="fa-solid fa-file-pdf"></i> Upload PDF Formula</h3>
                        <form action="/admin/upload-product-pdf" method="POST" enctype="multipart/form-data" style="margin-top: 15px;">
                            <div style="margin-bottom: 12px;">
                                <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Segmen:</label>
                                <select name="segment" required style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                                    ${segments.map(s => `<option value="${s}">${s}</option>`).join('')}
                                </select>
                            </div>
                            <div style="margin-bottom: 12px;">
                                <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Nama Produk:</label>
                                <input type="text" name="product_name" required placeholder="Nama Produk" style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                            </div>
                            <div style="margin-bottom: 12px;">
                                <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">File PDF:</label>
                                <input type="file" name="formula_pdf" accept=".pdf" required style="padding: 6px; background: #f8f9fa; border: 1px dashed #cbd5e0; border-radius: 6px; width: 100%; box-sizing: border-box;">
                            </div>
                            <button type="submit" style="background: #3182ce; color: #fff; border: none; padding: 9px 16px; font-weight: 600; border-radius: 6px; cursor: pointer; font-size: 0.9rem;">Upload PDF</button>
                        </form>
                    </div>
                </div>

                <div class="section-box" style="margin-top: 25px;">
                    <h3 style="color: #2d3748; margin-top: 0; margin-bottom: 15px;"><i class="fa-solid fa-layer-group"></i> Pilih Segmen Produk</h3>
                    <div>
                        ${segments.map(seg => `
                            <a href="/admin/products/segment/${encodeURIComponent(seg)}" class="segment-btn">
                                <span><i class="fa-solid fa-folder-open" style="color: #008744; margin-right: 10px;"></i> ${seg}</span>
                                <i class="fa-solid fa-chevron-right" style="color: #a0aec0; font-size: 0.9rem;"></i>
                            </a>
                        `).join('')}
                    </div>
                </div>
            </div>
        </body>
        </html>
    `);
});

// ==========================================
// HALAMAN DETAIL SEGMEN
// ==========================================
app.get('/admin/products/segment/:segmentName', async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/signin');

    const seg = decodeURIComponent(req.params.segmentName);
    const segIdSafe = seg.toLowerCase().replace(/[^a-z0-9]/g, '-');

    const prodRes = await db.execute({
        sql: `SELECT id, segment, principal, product_name, inci_name FROM products WHERE segment = ? ORDER BY principal, id DESC`,
        args: [seg]
    });
    const segProducts = prodRes.rows;

    const formulaRes = await db.execute({
        sql: `SELECT id, product_name, segment, pdf_url, created_at FROM product_formulas WHERE segment = ? ORDER BY created_at DESC`,
        args: [seg]
    });
    const segFormulas = formulaRes.rows;

    const segments = [
        "Personal Care, Cosmetic & Fragrance",
        "Pharmaceutical & Nutraceutical",
        "Home Care I&I",
        "Feedmill & Animal Health",
        "Processing Chemicals"
    ];

    let rowsHTML = segProducts.length > 0 ? segProducts.map(p => `
        <tr>
            <td style="text-align: center;"><input type="checkbox" name="selected_ids" value="${p.id}" class="product-checkbox-${segIdSafe}" onchange="updateBulkBtn('${segIdSafe}')"></td>
            <td><strong>${p.principal || '-'}</strong></td>
            <td>${p.product_name}</td>
            <td>${p.inci_name || '-'}</td>
            <td style="text-align: center;">
                <button type="button" onclick="openEditProductModal(${p.id}, '${String(p.principal || '').replace(/'/g, "\\'")}', '${String(p.product_name).replace(/'/g, "\\'")}', '${String(p.inci_name || '').replace(/'/g, "\\'")}', '${p.segment}')" style="background:#3182ce; color:#fff; border:none; padding:6px 12px; border-radius:6px; cursor:pointer; font-weight:600; font-size:0.85rem; margin-right:4px;">
                    <i class="fa-solid fa-pen-to-square"></i> Edit
                </button>
                <button type="button" onclick="openSingleDeleteModal(${p.id}, '${String(p.product_name).replace(/'/g, "\\'")}')" style="background:#e53e3e; color:#fff; border:none; padding:6px 12px; border-radius:6px; cursor:pointer; font-weight:600; font-size:0.85rem;">
                    <i class="fa-solid fa-trash"></i> Hapus
                </button>
            </td>
        </tr>
    `).join('') : `<tr><td colspan="5" style="text-align:center; padding: 25px; color: #718096; font-style: italic;">Belum ada produk pada segmen ini.</td></tr>`;

    let formulaRowsHTML = segFormulas.length > 0 ? segFormulas.map(f => `
        <tr>
            <td>${f.created_at}</td>
            <td><strong>${f.product_name}</strong></td>
            <td><a href="${f.pdf_url}" target="_blank" style="color: #3182ce; text-decoration: underline;"><i class="fa-solid fa-file-pdf"></i> Lihat PDF</a></td>
            <td style="text-align: center;">
                <button type="button" onclick="openEditFormulaModal(${f.id}, '${String(f.product_name).replace(/'/g, "\\'")}', '${f.segment}')" style="background:#3182ce; color:#fff; border:none; padding:6px 12px; border-radius:6px; cursor:pointer; font-weight:600; font-size:0.85rem; margin-right:4px;">
                    <i class="fa-solid fa-pen-to-square"></i> Edit
                </button>
                <button type="button" onclick="openDeleteFormulaModal(${f.id}, '${String(f.product_name).replace(/'/g, "\\'")}')" style="background:#e53e3e; color:#fff; border:none; padding:6px 12px; border-radius:6px; cursor:pointer; font-weight:600; font-size:0.85rem;">
                    <i class="fa-solid fa-trash"></i> Hapus
                </button>
            </td>
        </tr>
    `).join('') : `<tr><td colspan="4" style="text-align:center; padding: 20px; color: #718096; font-style: italic;">Belum ada riwayat upload PDF formula untuk segmen ini.</td></tr>`;

    res.send(`
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Segmen: ${seg} - PT Tritunggal Arthamakmur</title>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Open Sans', sans-serif; background: #f4f6f9; margin: 0; padding: 40px; color: #333; box-sizing: border-box; }
                .container { max-width: 1100px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
                h1 { color: #008744; margin-top: 0; font-size: 1.6rem; display: flex; align-items: center; gap: 10px; }
                .section-box { background: #fff; padding: 25px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; margin-top: 25px; }
                .table-responsive { width: 100%; overflow-x: auto; margin-top: 5px; }
                table { width: 100%; border-collapse: collapse; white-space: nowrap; }
                th, td { padding: 10px 12px; border: 1px solid #e2e8f0; text-align: left; font-size: 0.95rem; }
                th { background-color: #f8f9fa; color: #2d3748; }
                .btn-back { display: inline-block; background: #718096; color: #fff; padding: 9px 18px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 0.9rem; margin-bottom: 20px; }
                .btn-back:hover { background: #4a5568; }
                
                .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.5); justify-content: center; align-items: center; z-index: 1000; }
                .modal-overlay.active { display: flex; }
                .modal-card { background: #fff; width: 100%; max-width: 420px; padding: 30px; border-radius: 12px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.15); box-sizing: border-box; }
                .modal-icon { font-size: 3rem; color: #e53e3e; margin-bottom: 15px; }
                .modal-card h3 { margin: 0 0 10px 0; font-size: 1.3rem; color: #1a202c; }
                .modal-card p { color: #4a5568; font-size: 0.95rem; margin-bottom: 25px; line-height: 1.5; }
                .modal-actions { display: flex; gap: 12px; justify-content: center; }
                .btn-modal-cancel { background: #e2e8f0; color: #2d3748; border: none; padding: 10px 20px; border-radius: 6px; font-weight: 600; cursor: pointer; flex: 1; }
                .btn-modal-confirm { background: #e53e3e; color: #fff; border: none; padding: 10px 20px; border-radius: 6px; font-weight: 600; cursor: pointer; flex: 1; }
            </style>
        </head>
        <body>
            <div class="container">
                <a href="/admin/products/dashboard" class="btn-back"><i class="fa-solid fa-arrow-left"></i> Kembali ke Menu Kelola Produk</a>
                <h1><i class="fa-solid fa-layer-group"></i> Segmen: ${seg}</h1>
                <p>Kelola riwayat PDF formula dan daftar produk untuk segmen ini.</p>

                <div class="section-box">
                    <h3 style="color: #2d3748; margin-top: 0; margin-bottom: 10px; display: flex; align-items: center; gap: 8px;">
                        <i class="fa-solid fa-clock-rotate-left"></i> Riwayat Upload PDF Formula (${segFormulas.length})
                    </h3>
                    <div class="table-responsive">
                        <table>
                            <thead>
                                <tr>
                                    <th>Waktu Upload</th>
                                    <th>Nama Produk</th>
                                    <th>File PDF</th>
                                    <th style="text-align: center; width: 150px;">Aksi</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${formulaRowsHTML}
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="section-box">
                    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 10px;">
                        <h3 style="color: #2d3748; margin: 0; display: flex; align-items: center; gap: 8px;">
                            <i class="fa-solid fa-boxes-stacked"></i> Daftar Produk (${segProducts.length})
                        </h3>
                        ${segProducts.length > 0 ? `
                            <button type="button" id="bulkBtn-${segIdSafe}" onclick="openBulkDeleteModal('${segIdSafe}')" style="background:#e53e3e; color:#fff; border:none; padding:7px 14px; border-radius:6px; cursor:pointer; font-weight:600; font-size: 0.9rem; display:none;">
                                <i class="fa-solid fa-trash"></i> Hapus Pilihan (<span id="count-${segIdSafe}">0</span>)
                            </button>
                        ` : ''}
                    </div>

                    <form id="bulkDeleteForm" action="/admin/products/bulk-delete" method="POST">
                        <div class="table-responsive">
                            <table>
                                <thead>
                                    <tr>
                                        <th style="width: 40px; text-align: center;">
                                            ${segProducts.length > 0 ? `<input type="checkbox" id="selectAll-${segIdSafe}" onclick="toggleSelectAll('${segIdSafe}', this)">` : ''}
                                        </th>
                                        <th>Principal</th>
                                        <th>Nama Produk</th>
                                        <th>INCI Name</th>
                                        <th style="text-align: center; width: 140px;">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${rowsHTML}
                                </tbody>
                            </table>
                        </div>
                    </form>
                </div>
            </div>

            <div id="customDeleteModal" class="modal-overlay">
                <div class="modal-card">
                    <div class="modal-icon"><i class="fa-solid fa-triangle-exclamation"></i></div>
                    <h3 id="modalTitle">Konfirmasi Hapus</h3>
                    <p id="modalDeleteMessage">Apakah Anda yakin ingin menghapus item ini?</p>
                    <div class="modal-actions">
                        <button type="button" class="btn-modal-cancel" onclick="closeCustomModal()">Batal</button>
                        <button type="button" class="btn-modal-confirm" id="modalConfirmBtn" onclick="executeDelete()">Ya, Hapus</button>
                    </div>
                </div>
            </div>

            <div id="editFormulaModal" class="modal-overlay">
                <div class="modal-card" style="max-width: 480px; text-align: left;">
                    <h3 style="margin-top: 0; color: #008744; display: flex; align-items: center; gap: 8px;">
                        <i class="fa-solid fa-pen-to-square"></i> Edit PDF Formula
                    </h3>
                    <form id="editFormulaForm" method="POST" enctype="multipart/form-data">
                        <div style="margin-bottom: 15px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Nama Produk:</label>
                            <input type="text" id="editFormulaProductName" name="product_name" required style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                        </div>
                        <div style="margin-bottom: 15px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Segmen:</label>
                            <select id="editFormulaSegment" name="segment" required style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                                ${segments.map(s => `<option value="${s}">${s}</option>`).join('')}
                            </select>
                        </div>
                        <div style="margin-bottom: 20px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Ganti File PDF Baru (Opsional):</label>
                            <input type="file" name="formula_pdf" accept=".pdf" style="padding: 6px; background: #f8f9fa; border: 1px dashed #cbd5e0; border-radius: 6px; width: 100%; box-sizing: border-box;">
                            <small style="color: #718096; display: block; margin-top: 4px;">Biarkan kosong jika hanya ingin mengubah nama/segmen produk.</small>
                        </div>
                        <div style="display: flex; gap: 10px; justify-content: flex-end;">
                            <button type="button" class="btn-modal-cancel" onclick="closeEditFormulaModal()" style="flex: none; padding: 8px 16px;">Batal</button>
                            <button type="submit" style="background: #008744; color: #fff; border: none; padding: 8px 20px; border-radius: 6px; font-weight: 600; cursor: pointer;">Simpan Perubahan</button>
                        </div>
                    </form>
                </div>
            </div>

            <div id="editProductModal" class="modal-overlay">
                <div class="modal-card" style="max-width: 480px; text-align: left;">
                    <h3 style="margin-top: 0; color: #008744; display: flex; align-items: center; gap: 8px;">
                        <i class="fa-solid fa-pen-to-square"></i> Edit Data Produk
                    </h3>
                    <form id="editProductForm" method="POST">
                        <div style="margin-bottom: 12px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Segmen:</label>
                            <select id="editProductSegment" name="segment" required style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                                ${segments.map(s => `<option value="${s}">${s}</option>`).join('')}
                            </select>
                        </div>
                        <div style="margin-bottom: 12px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Principal:</label>
                            <input type="text" id="editProductPrincipal" name="principal" placeholder="Principal" style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                        </div>
                        <div style="margin-bottom: 12px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">Nama Produk:</label>
                            <input type="text" id="editProductName" name="product_name" required placeholder="Nama Produk" style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                        </div>
                        <div style="margin-bottom: 20px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px; font-size:0.9rem;">INCI Name:</label>
                            <input type="text" id="editProductInci" name="inci_name" placeholder="INCI Name" style="width: 100%; padding: 8px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box;">
                        </div>
                        <div style="display: flex; gap: 10px; justify-content: flex-end;">
                            <button type="button" class="btn-modal-cancel" onclick="closeEditProductModal()" style="flex: none; padding: 8px 16px;">Batal</button>
                            <button type="submit" style="background: #008744; color: #fff; border: none; padding: 8px 20px; border-radius: 6px; font-weight: 600; cursor: pointer;">Simpan Perubahan</button>
                        </div>
                    </form>
                </div>
            </div>

            <form id="singleDeleteForm" method="POST" style="display: none;"></form>

            <script>
                let deleteActionType = '';
                let targetId = null;
                let activeSegId = null;

                function toggleSelectAll(segId, source) {
                    const checkboxes = document.querySelectorAll('.product-checkbox-' + segId);
                    checkboxes.forEach(cb => cb.checked = source.checked);
                    updateBulkBtn(segId);
                }

                function updateBulkBtn(segId) {
                    const checkedCount = document.querySelectorAll('.product-checkbox-' + segId + ':checked').length;
                    const btn = document.getElementById('bulkBtn-' + segId);
                    const countSpan = document.getElementById('count-' + segId);
                    
                    if (countSpan) countSpan.textContent = checkedCount;
                    if (btn) {
                        btn.style.display = checkedCount > 0 ? 'inline-block' : 'none';
                    }
                }

                function openSingleDeleteModal(id, productName) {
                    deleteActionType = 'single';
                    targetId = id;
                    document.getElementById('modalTitle').textContent = 'Konfirmasi Hapus Produk';
                    document.getElementById('modalDeleteMessage').innerHTML = 'Apakah Anda yakin ingin menghapus produk <strong>"' + productName + '"</strong>?';
                    document.getElementById('customDeleteModal').classList.add('active');
                }

                function openDeleteFormulaModal(id, productName) {
                    deleteActionType = 'formula';
                    targetId = id;
                    document.getElementById('modalTitle').textContent = 'Konfirmasi Hapus Riwayat PDF';
                    document.getElementById('modalDeleteMessage').innerHTML = 'Apakah Anda yakin ingin menghapus riwayat PDF formula untuk produk <strong>"' + productName + '"</strong>?';
                    document.getElementById('customDeleteModal').classList.add('active');
                }

                function openBulkDeleteModal(segId) {
                    deleteActionType = 'bulk';
                    activeSegId = segId;
                    const count = document.querySelectorAll('.product-checkbox-' + segId + ':checked').length;
                    document.getElementById('modalTitle').textContent = 'Konfirmasi Hapus Massal';
                    document.getElementById('modalDeleteMessage').innerHTML = 'Apakah Anda yakin ingin menghapus <strong>' + count + ' produk</strong> yang dipilih pada segmen ini?';
                    document.getElementById('customDeleteModal').classList.add('active');
                }

                function closeCustomModal() {
                    document.getElementById('customDeleteModal').classList.remove('active');
                }

                function executeDelete() {
                    const form = document.getElementById('singleDeleteForm');
                    if (deleteActionType === 'single') {
                        form.action = '/admin/products/delete/' + targetId;
                        form.submit();
                    } else if (deleteActionType === 'formula') {
                        form.action = '/admin/products/delete-formula/' + targetId;
                        form.submit();
                    } else if (deleteActionType === 'bulk') {
                        document.getElementById('bulkDeleteForm').submit();
                    }
                }

                function openEditFormulaModal(id, productName, segment) {
                    document.getElementById('editFormulaForm').action = '/admin/products/edit-formula/' + id;
                    document.getElementById('editFormulaProductName').value = productName;
                    document.getElementById('editFormulaSegment').value = segment;
                    document.getElementById('editFormulaModal').classList.add('active');
                }

                function closeEditFormulaModal() {
                    document.getElementById('editFormulaModal').classList.remove('active');
                }

                function openEditProductModal(id, principal, productName, inciName, segment) {
                    document.getElementById('editProductForm').action = '/admin/products/edit/' + id;
                    document.getElementById('editProductPrincipal').value = principal;
                    document.getElementById('editProductName').value = productName;
                    document.getElementById('editProductInci').value = inciName;
                    document.getElementById('editProductSegment').value = segment;
                    document.getElementById('editProductModal').classList.add('active');
                }

                function closeEditProductModal() {
                    document.getElementById('editProductModal').classList.remove('active');
                }
            </script>
        </body>
        </html>
    `);
});

app.post('/admin/products/edit/:id', async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    const productId = req.params.id;
    const { segment, principal, product_name, inci_name } = req.body;

    try {
        await db.execute({
            sql: `UPDATE products SET segment = ?, principal = ?, product_name = ?, inci_name = ? WHERE id = ?`,
            args: [segment, principal, product_name, inci_name, productId]
        });
        res.redirect('/admin/products/dashboard');
    } catch (err) {
        res.status(500).send('Gagal memperbarui data produk: ' + err.message);
    }
});

app.post('/admin/products/edit-formula/:id', upload.single('formula_pdf'), async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    const formulaId = req.params.id;
    const { product_name, segment } = req.body;

    try {
        const oldRes = await db.execute({
            sql: `SELECT * FROM product_formulas WHERE id = ?`,
            args: [formulaId]
        });
        const oldFormula = oldRes.rows[0];

        if (!oldFormula) {
            return res.status(404).send('Data riwayat formula tidak ditemukan.');
        }

        let pdfUrl = oldFormula.pdf_url;
        if (req.file) {
            pdfUrl = `/uploads/${req.file.filename}`;
        }

        await db.execute({
            sql: `UPDATE product_formulas SET product_name = ?, segment = ?, pdf_url = ? WHERE id = ?`,
            args: [product_name, segment, pdfUrl, formulaId]
        });

        await db.execute({
            sql: `UPDATE products SET formula_pdf = ? WHERE product_name = ? AND segment = ?`,
            args: [pdfUrl, product_name, segment]
        });

        res.redirect('/admin/products/dashboard');
    } catch (err) {
        res.status(500).send('Gagal memperbarui riwayat PDF: ' + err.message);
    }
});

app.post('/admin/products/delete-formula/:id', async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    const formulaId = req.params.id;

    try {
        const formulaRes = await db.execute({
            sql: `SELECT product_name, segment, pdf_url FROM product_formulas WHERE id = ?`,
            args: [formulaId]
        });
        const formula = formulaRes.rows[0];

        if (formula) {
            await db.execute({
                sql: `UPDATE products SET formula_pdf = NULL WHERE product_name = ? AND segment = ? AND formula_pdf = ?`,
                args: [formula.product_name, formula.segment, formula.pdf_url]
            });
        }

        await db.execute({
            sql: `DELETE FROM product_formulas WHERE id = ?`,
            args: [formulaId]
        });

        res.redirect('/admin/products/dashboard');
    } catch (err) {
        res.status(500).send('Gagal menghapus riwayat PDF: ' + err.message);
    }
});

app.post('/admin/products/delete/:id', async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    try {
        await db.execute({
            sql: `DELETE FROM products WHERE id = ?`,
            args: [req.params.id]
        });
        res.redirect('/admin/products/dashboard');
    } catch (err) {
        res.status(500).send('Gagal menghapus produk: ' + err.message);
    }
});

app.post('/admin/products/bulk-delete', async (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    try {
        let { selected_ids } = req.body;
        if (!selected_ids) {
            return res.status(400).send('Tidak ada produk yang dipilih.');
        }

        if (!Array.isArray(selected_ids)) {
            selected_ids = [selected_ids];
        }

        const placeholders = selected_ids.map(() => '?').join(',');
        await db.execute({
            sql: `DELETE FROM products WHERE id IN (${placeholders})`,
            args: selected_ids
        });

        res.redirect('/admin/products/dashboard');
    } catch (err) {
        res.status(500).send('Gagal menghapus produk massal: ' + err.message);
    }
});

// ==========================================
// ADMIN CMS MANAGEMENT ENDPOINTS
// ==========================================
app.get('/admin/cms', async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/signin');

    const settingsRes = await db.execute(`SELECT * FROM site_settings`);
    let settings = {};
    settingsRes.rows.forEach(row => { settings[row.key] = row.value; });

    const partnerRes = await db.execute(`SELECT * FROM partners ORDER BY id DESC`);
    const partners = partnerRes.rows;

    const transRes = await db.execute(`SELECT * FROM translations ORDER BY id ASC`);
    const transList = transRes.rows;

    let transRows = transList.map(tr => `
        <tr>
            <td><code>${tr.translation_key}</code></td>
            <td><input type="text" name="en_${tr.translation_key}" value="${String(tr.en_text).replace(/"/g, '&quot;')}" style="width:100%; padding:6px; box-sizing:border-box;"></td>
            <td><input type="text" name="id_${tr.translation_key}" value="${String(tr.id_text).replace(/"/g, '&quot;')}" style="width:100%; padding:6px; box-sizing:border-box;"></td>
        </tr>
    `).join('');

    let partnerRows = partners.map(p => `
        <tr>
            <td style="text-align: center;"><img src="${p.logo_url}" style="height: 35px; object-fit: contain;"></td>
            <td><strong>${p.name}</strong></td>
            <td><a href="${p.website_url || '#'}" target="_blank">${p.website_url || '-'}</a></td>
            <td style="text-align: center;">
                <button type="button" 
                        onclick="openEditPartnerModal(${p.id}, '${String(p.name).replace(/'/g, "\\'")}', '${String(p.website_url || '').replace(/'/g, "\\'")}', '${p.logo_url}')" 
                        style="background:#3182ce; color:#fff; border:none; padding:6px 12px; border-radius:4px; cursor:pointer; font-weight:600; margin-right:5px;">
                    Edit
                </button>
                <form action="/admin/partners/delete/${p.id}" method="POST" style="display:inline;">
                    <button type="submit" onclick="return confirm('Yakin ingin menghapus partner ini?')" style="background:#e53e3e; color:#fff; border:none; padding:6px 12px; border-radius:4px; cursor:pointer; font-weight:600;">Hapus</button>
                </form>
            </td>
        </tr>
    `).join('');

    res.send(`
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Admin CMS Management - PT Tritunggal Arthamakmur</title>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Open Sans', sans-serif; background: #f4f6f9; margin: 0; padding: 40px; color: #333; box-sizing: border-box; }
                .container { max-width: 1000px; margin: 0 auto; }
                .header-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; background: #fff; padding: 20px 25px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); gap: 15px; flex-wrap: wrap; }
                .header-title-wrapper { display: flex; align-items: center; gap: 12px; min-width: 0; flex: 1; flex-wrap: wrap; }
                .header-bar h1 { color: #008744; margin: 0; font-size: 1.35rem; font-weight: 700; }
                .header-bar i.fa-pen-to-square { font-size: 1.5rem; color: #008744; flex-shrink: 0; }
                .btn-back { background: #718096; color: #fff; padding: 9px 16px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 0.9rem; white-space: nowrap; flex-shrink: 0; transition: background 0.2s; }
                .btn-back:hover { background: #4a5568; }
                .section-card { background: #fff; padding: 30px; border-radius: 10px; margin-bottom: 30px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); box-sizing: border-box; width: 100%; overflow: hidden; }
                .section-card h2 { color: #1a202c; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; margin-top: 0; margin-bottom: 20px; font-size: 1.3rem; display: flex; align-items: center; gap: 10px; }
                .section-card h2 i { color: #008744; }
                .form-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 20px; }
                .form-group { margin-bottom: 18px; min-width: 0; }
                .form-group.full-width { grid-column: 1 / -1; }
                label { display: block; font-weight: 600; margin-bottom: 6px; color: #2d3748; font-size: 0.95rem; }
                input[type="text"], input[type="url"], textarea { width: 100%; padding: 10px 14px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box; font-family: inherit; font-size: 0.95rem; }
                input[type="file"] { display: block; width: 100%; padding: 8px; background: #f8f9fa; border: 1px dashed #cbd5e0; border-radius: 6px; cursor: pointer; box-sizing: border-box; }
                textarea { height: 90px; resize: vertical; }
                .btn-save { background: #008744; color: #fff; border: none; padding: 12px 28px; font-weight: 700; border-radius: 6px; cursor: pointer; font-size: 1rem; }
                .btn-save:hover { background: #006b35; }
                .table-responsive { width: 100%; overflow-x: auto; margin-top: 15px; }
                table { width: 100%; border-collapse: collapse; white-space: nowrap; }
                th, td { border: 1px solid #e2e8f0; padding: 12px; text-align: left; font-size: 0.95rem; }
                th { background-color: #f8f9fa; }
                .preview-img { max-height: 40px; margin-top: 8px; border: 1px solid #cbd5e0; padding: 4px; border-radius: 4px; background: #f7fafc; max-width: 100%; object-fit: contain; }
                .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); justify-content: center; align-items: center; z-index: 1000; padding: 15px; box-sizing: border-box; }
                .modal-overlay.active { display: flex; }
                .modal-box { background: #fff; width: 100%; max-width: 500px; padding: 25px; border-radius: 8px; box-sizing: border-box; }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header-bar">
                    <div class="header-title-wrapper">
                        <i class="fa-solid fa-pen-to-square"></i>
                        <h1>Kelola Konten Website (CMS)</h1>
                    </div>
                    <a href="/admin/dashboard" class="btn-back">&larr; Kembali ke Dashboard</a>
                </div>

                <form action="/admin/settings/update" method="POST" enctype="multipart/form-data">
                    <div class="section-card">
                        <h2><i class="fa-solid fa-compass"></i> Bagian 1: Navbar & Logo Header</h2>
                        <div class="form-grid">
                            <div class="form-group">
                                <label>Logo Utama Awal</label>
                                <input type="file" name="logo_white_file" accept="image/*">
                                ${settings.logo_white ? `<img src="${settings.logo_white}" class="preview-img" style="background:#333;">` : ''}
                            </div>
                            <div class="form-group">
                                <label>Logo Setelah Scrolled</label>
                                <input type="file" name="logo_dark_file" accept="image/*">
                                ${settings.logo_dark ? `<img src="${settings.logo_dark}" class="preview-img">` : ''}
                            </div>
                        </div>
                    </div>

                    <div class="section-card">
                        <h2><i class="fa-solid fa-window-maximize"></i> Bagian 2: Body & Konten Halaman Utama</h2>
                        <div class="form-grid">
                            <div class="form-group">
                                <label>Judul Section (Homepage)</label>
                                <input type="text" name="who_we_are_title" value="${settings.who_we_are_title || 'Who We Are'}">
                            </div>
                            <div class="form-group full-width">
                                <label>Deskripsi Section (Homepage)</label>
                                <textarea name="who_we_are_text" rows="4">${settings.who_we_are_text || ''}</textarea>
                            </div>
                        </div>
                        <div class="form-grid" style="margin-top: 15px;">
                            <div class="form-group full-width">
                                <label>Judul Banner Bawah (Homepage)</label>
                                <input type="text" name="cta_title" value="${settings.cta_title || ''}">
                            </div>
                            <div class="form-group full-width">
                                <label>Teks Deskripsi Banner Bawah (Homepage)</label>
                                <textarea name="cta_text" rows="2">${settings.cta_text || ''}</textarea>
                            </div>
                            <div class="form-group full-width">
                                <label>Teks Deskripsi Hero About Us</label>
                                <textarea name="about_hero_text" rows="3">${settings.about_hero_text || ''}</textarea>
                            </div>
                            <div class="form-group full-width">
                                <label>Teks Intro / Deskripsi Halaman News (Area Putih)</label>
                                <textarea name="news_hero_text" rows="3">${settings.news_hero_text || ''}</textarea>
                            </div>
                            <div class="form-group full-width">
                                <label>Teks Deskripsi / Subtitle Bawah "Let us know what you need"</label>
                                <textarea name="contact_subtext" rows="2">${settings.contact_subtext || ''}</textarea>
                            </div>
                        </div>
                        
                        <div class="section-card" style="margin-top: 25px;">
                            <h2><i class="fa-solid fa-images"></i> Bagian Beranda: Slider Background Hero (5 Gambar)</h2>
                            <div class="form-grid">
                                <div class="form-group">
                                    <label>Slider Background 1</label>
                                    <input type="file" name="hero_bg_file_1" accept="image/*">
                                    ${settings.hero_bg_image_1 ? `<img src="${settings.hero_bg_image_1}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Slider Background 2</label>
                                    <input type="file" name="hero_bg_file_2" accept="image/*">
                                    ${settings.hero_bg_image_2 ? `<img src="${settings.hero_bg_image_2}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Slider Background 3</label>
                                    <input type="file" name="hero_bg_file_3" accept="image/*">
                                    ${settings.hero_bg_image_3 ? `<img src="${settings.hero_bg_image_3}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Slider Background 4</label>
                                    <input type="file" name="hero_bg_file_4" accept="image/*">
                                    ${settings.hero_bg_image_4 ? `<img src="${settings.hero_bg_image_4}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Slider Background 5</label>
                                    <input type="file" name="hero_bg_file_5" accept="image/*">
                                    ${settings.hero_bg_image_5 ? `<img src="${settings.hero_bg_image_5}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                            </div>
                        </div>

                        <div class="form-group full-width">
                            <label>Gambar Banner Bawah (Homepage)</label>
                            <input type="file" name="cta_bg_file" accept="image/*">
                            ${settings.cta_bg_image ? `<img src="${settings.cta_bg_image}" class="preview-img" style="max-height: 80px;">` : ''}
                        </div>
                        <div class="form-group">
                            <label>Gambar Latar Belakang (Contact Us)</label>
                            <input type="file" name="contact_bg_file" accept="image/*">
                            ${settings.contact_bg_image ? `<img src="${settings.contact_bg_image}" class="preview-img" style="max-height: 80px;">` : ''}
                        </div>
                        <div class="form-group">
                            <label>Gambar Latar Belakang (News)</label>
                            <input type="file" name="news_bg_file" accept="image/*">
                            ${settings.news_bg_image ? `<img src="${settings.news_bg_image}" class="preview-img" style="max-height: 80px;">` : ''}
                        </div>
                        <div class="form-group full-width">
                            <label>Gambar Banner Bawah (Newsletter)</label>
                            <input type="file" name="newsletter_bg_file" accept="image/*">
                            ${settings.newsletter_bg_image ? `<img src="${settings.newsletter_bg_image}" class="preview-img" style="max-height: 80px;">` : ''}
                        </div>

                        <div class="section-card">
                            <h2><i class="fa-solid fa-handshake"></i> Kelola Halaman Principals</h2>
                            <div class="form-grid">
                                <div class="form-group full-width">
                                    <label>Gambar Latar Belakang (Principals Page)</label>
                                    <input type="file" name="principals_bg_file" accept="image/*">
                                    ${settings.principals_bg_image ? `<img src="${settings.principals_bg_image}" class="preview-img" style="max-height: 80px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Teks Pengantar / Intro Principals</label>
                                    <textarea name="principals_intro_text" rows="3">${settings.principals_intro_text || ''}</textarea>
                                </div>
                            </div>
                        </div>

                        <div class="section-card">
                            <h2><i class="fa-solid fa-address-card"></i> Kelola Halaman About Us (Banner & 4 Tab Gambar)</h2>
                            <div class="form-grid">
                                <div class="form-group full-width">
                                    <label>Banner Background (Hero About Us)</label>
                                    <input type="file" name="about_bg_file" accept="image/*">
                                    ${settings.about_bg_image ? `<img src="${settings.about_bg_image}" class="preview-img" style="max-height: 80px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Gambar Tab 1: About Us</label>
                                    <input type="file" name="about_tab_file_1" accept="image/*">
                                    ${settings.about_tab_img_1 ? `<img src="${settings.about_tab_img_1}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Gambar Tab 2: Historical Highlights</label>
                                    <input type="file" name="about_tab_file_2" accept="image/*">
                                    ${settings.about_tab_img_2 ? `<img src="${settings.about_tab_img_2}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Gambar Tab 3: Vision & Mission</label>
                                    <input type="file" name="about_tab_file_3" accept="image/*">
                                    ${settings.about_tab_img_3 ? `<img src="${settings.about_tab_img_3}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Gambar Tab 4: Strategic Presence</label>
                                    <input type="file" name="about_tab_file_4" accept="image/*">
                                    ${settings.about_tab_img_4 ? `<img src="${settings.about_tab_img_4}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                            </div>
                        </div>

                        <div class="section-card" style="margin-top: 30px; border: 1px solid #e2e8f0;">
                            <h2><i class="fa-solid fa-layer-group"></i> Kelola Banner & Aplikasi Segmen Produk</h2>
                            
                            <h4 style="margin-top:15px; color:#008744;">1. Personal Care, Cosmetic & Fragrance</h4>
                            <div class="form-grid">
                                <div class="form-group">
                                    <label>Judul Segment 1</label>
                                    <input type="text" name="segment_personal_care_title" value="${settings.segment_personal_care_title || ''}">
                                </div>  
                                <div class="form-group">
                                    <label>Banner Segment 1</label>
                                    <input type="file" name="segment_personal_care_bg_file" accept="image/*">
                                    ${settings.segment_personal_care_bg ? `<img src="${settings.segment_personal_care_bg}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi / Paragraf Intro Segment 1</label>
                                    <textarea name="segment_personal_care_intro" rows="3">${settings.segment_personal_care_intro || ''}</textarea>
                                </div>
                                <div class="form-group">
                                    <label>Judul App 1</label>
                                    <input type="text" name="app_pc_cosmetics_title" value="${settings.app_pc_cosmetics_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 1</label>
                                    <input type="file" name="app_pc_cosmetics_img_file" accept="image/*">
                                    ${settings.app_pc_cosmetics_img ? `<img src="${settings.app_pc_cosmetics_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Judul App 2</label>
                                    <input type="text" name="app_pc_skincare_title" value="${settings.app_pc_skincare_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 2</label>
                                    <input type="file" name="app_pc_skincare_img_file" accept="image/*">
                                    ${settings.app_pc_skincare_img ? `<img src="${settings.app_pc_skincare_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Judul App 3</label>
                                    <input type="text" name="app_pc_bath_title" value="${settings.app_pc_bath_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 3</label>
                                    <input type="file" name="app_pc_bath_img_file" accept="image/*">
                                    ${settings.app_pc_bath_img ? `<img src="${settings.app_pc_bath_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Judul App 4</label>
                                    <input type="text" name="app_pc_haircare_title" value="${settings.app_pc_haircare_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 4</label>
                                    <input type="file" name="app_pc_haircare_img_file" accept="image/*">
                                    ${settings.app_pc_haircare_img ? `<img src="${settings.app_pc_haircare_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                            </div>

                            <hr style="border: 0; border-top: 1px dashed #cbd5e0; margin: 20px 0;">

                            <h4 style="color:#008744;">2. Pharmaceutical & Nutraceutical</h4>
                            <div class="form-grid">
                                <div class="form-group">
                                    <label>Judul Segment 2</label>
                                    <input type="text" name="segment_pharma_title" value="${settings.segment_pharma_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Banner Segment 2</label>
                                    <input type="file" name="segment_pharma_bg_file" accept="image/*">
                                    ${settings.segment_pharma_bg ? `<img src="${settings.segment_pharma_bg}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi / Paragraf Intro Segment 2</label>
                                    <textarea name="segment_pharma_intro" rows="3">${settings.segment_pharma_intro || ''}</textarea>
                                </div>
                                <div class="form-group">
                                    <label>Judul App 1</label>
                                    <input type="text" name="app_pharma_pharma_title" value="${settings.app_pharma_pharma_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 1</label>
                                    <input type="file" name="app_pharma_pharma_img_file" accept="image/*">
                                    ${settings.app_pharma_pharma_img ? `<img src="${settings.app_pharma_pharma_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Judul App 2</label>
                                    <input type="text" name="app_pharma_food_title" value="${settings.app_pharma_food_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 2</label>
                                    <input type="file" name="app_pharma_food_img_file" accept="image/*">
                                    ${settings.app_pharma_food_img ? `<img src="${settings.app_pharma_food_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Judul App 3</label>
                                    <input type="text" name="app_pharma_excipient_title" value="${settings.app_pharma_excipient_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 3</label>
                                    <input type="file" name="app_pharma_excipient_img_file" accept="image/*">
                                    ${settings.app_pharma_excipient_img ? `<img src="${settings.app_pharma_excipient_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                            </div>

                            <hr style="border: 0; border-top: 1px dashed #cbd5e0; margin: 20px 0;">

                            <h4 style="color:#008744;">3. Home Care I&I</h4>
                            <div class="form-grid">
                                <div class="form-group">
                                    <label>Judul Segment 3</label>
                                    <input type="text" name="segment_homecare_title" value="${settings.segment_homecare_title || ''}">
                                </div> 
                                <div class="form-group">
                                    <label>Banner Segment 3</label>
                                    <input type="file" name="segment_homecare_bg_file" accept="image/*">
                                    ${settings.segment_homecare_bg ? `<img src="${settings.segment_homecare_bg}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi / Paragraf Intro Segment 3</label>
                                    <textarea name="segment_homecare_intro" rows="3">${settings.segment_homecare_intro || ''}</textarea>
                                </div>
                                <div class="form-group">
                                    <label>Judul App 1</label>
                                    <input type="text" name="app_homecare_home_title" value="${settings.app_homecare_home_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 1</label>
                                    <input type="file" name="app_homecare_home_img_file" accept="image/*">
                                    ${settings.app_homecare_home_img ? `<img src="${settings.app_homecare_home_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Judul App 2</label>
                                    <input type="text" name="app_homecare_inst_title" value="${settings.app_homecare_inst_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 2</label>
                                    <input type="file" name="app_homecare_inst_img_file" accept="image/*">
                                    ${settings.app_homecare_inst_img ? `<img src="${settings.app_homecare_inst_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Judul App 3</label>
                                    <input type="text" name="app_homecare_ind_title" value="${settings.app_homecare_ind_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 3</label>
                                    <input type="file" name="app_homecare_ind_img_file" accept="image/*">
                                    ${settings.app_homecare_ind_img ? `<img src="${settings.app_homecare_ind_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                            </div>

                            <hr style="border: 0; border-top: 1px dashed #cbd5e0; margin: 20px 0;">

                            <h4 style="color:#008744;">4. Feedmill & Animal Health</h4>
                            <div class="form-grid">
                                <div class="form-group">
                                    <label>Judul Segment 4</label>
                                    <input type="text" name="segment_vet_title" value="${settings.segment_vet_title || ''}">
                                </div> 
                                <div class="form-group">
                                    <label>Banner Segment 4</label>
                                    <input type="file" name="segment_vet_bg_file" accept="image/*">
                                    ${settings.segment_vet_bg ? `<img src="${settings.segment_vet_bg}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi / Paragraf Intro Segment 4</label>
                                    <textarea name="segment_vet_intro" rows="3">${settings.segment_vet_intro || ''}</textarea>
                                </div>
                                <div class="form-group">
                                    <label>Judul App 1</label>
                                    <input type="text" name="app_vet_aqua_title" value="${settings.app_vet_aqua_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 1</label>
                                    <input type="file" name="app_vet_aqua_img_file" accept="image/*">
                                    ${settings.app_vet_aqua_img ? `<img src="${settings.app_vet_aqua_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group">
                                    <label>Judul App 2</label>
                                    <input type="text" name="app_vet_livestock_title" value="${settings.app_vet_livestock_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 2</label>
                                    <input type="file" name="app_vet_livestock_img_file" accept="image/*">
                                    ${settings.app_vet_livestock_img ? `<img src="${settings.app_vet_livestock_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                            </div>

                            <hr style="border: 0; border-top: 1px dashed #cbd5e0; margin: 20px 0;">

                            <h4 style="color:#008744;">5. Processing Chemicals</h4>
                            <div class="form-grid">
                                <div class="form-group">
                                    <label>Judul Segment 5</label>
                                    <input type="text" name="segment_processing_title" value="${settings.segment_processing_title || ''}">
                                </div> 
                                <div class="form-group full-width">
                                    <label>Deskripsi / Paragraf Intro Segmen Processing Chemicals (Area Putih)</label>
                                    <textarea name="segment_processing_intro" rows="3">${settings.segment_processing_intro || ''}</textarea>
                                </div>
                                <div class="form-group">
                                    <label>Banner Segment 5</label>
                                    <input type="file" name="segment_processing_bg_file" accept="image/*">
                                    ${settings.segment_processing_bg ? `<img src="${settings.segment_processing_bg}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>

                                <div class="form-group">
                                    <label>Judul App 1</label>
                                    <input type="text" name="app_proc_metalworking_title" value="${settings.app_proc_metalworking_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 1</label>
                                    <input type="file" name="app_proc_metalworking_img_file" accept="image/*">
                                    ${settings.app_proc_metalworking_img ? `<img src="${settings.app_proc_metalworking_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi App 1</label>
                                    <textarea name="processing_metalworking_desc" rows="2">${settings.processing_metalworking_desc || ''}</textarea>
                                </div>

                                <div class="form-group">
                                    <label>Judul App 2</label>
                                    <input type="text" name="app_proc_battery_title" value="${settings.app_proc_battery_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 2</label>
                                    <input type="file" name="app_proc_battery_img_file" accept="image/*">
                                    ${settings.app_proc_battery_img ? `<img src="${settings.app_proc_battery_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi App 2</label>
                                    <textarea name="processing_battery_desc" rows="2">${settings.processing_battery_desc || ''}</textarea>
                                </div>

                                <div class="form-group">
                                    <label>Judul App 3</label>
                                    <input type="text" name="app_proc_textile_title" value="${settings.app_proc_textile_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 3</label>
                                    <input type="file" name="app_proc_textile_img_file" accept="image/*">
                                    ${settings.app_proc_textile_img ? `<img src="${settings.app_proc_textile_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi App 3</label>
                                    <textarea name="processing_textile_desc" rows="2">${settings.processing_textile_desc || ''}</textarea>
                                </div>

                                <div class="form-group">
                                    <label>Judul App 4</label>
                                    <input type="text" name="app_proc_water_title" value="${settings.app_proc_water_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 4</label>
                                    <input type="file" name="app_proc_water_img_file" accept="image/*">
                                    ${settings.app_proc_water_img ? `<img src="${settings.app_proc_water_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi App 4</label>
                                    <textarea name="processing_water_desc" rows="2">${settings.processing_water_desc || ''}</textarea>
                                </div>

                                <div class="form-group">
                                    <label>Judul App 5</label>
                                    <input type="text" name="app_proc_cropcare_title" value="${settings.app_proc_cropcare_title || ''}">
                                </div>
                                <div class="form-group">
                                    <label>Gambar App 5</label>
                                    <input type="file" name="app_proc_cropcare_img_file" accept="image/*">
                                    ${settings.app_proc_cropcare_img ? `<img src="${settings.app_proc_cropcare_img}" class="preview-img" style="max-height: 60px;">` : ''}
                                </div>
                                <div class="form-group full-width">
                                    <label>Deskripsi App 5</label>
                                    <textarea name="processing_cropcare_desc" rows="2">${settings.processing_cropcare_desc || ''}</textarea>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="section-card">
                        <h2><i class="fa-solid fa-shoe-prints"></i> Bagian 3: Footer Lengkap</h2>
                        <div class="form-grid">
                            <div class="form-group full-width">
                                <label>Nama Perusahaan Footer</label>
                                <input type="text" name="footer_company_name" value="${settings.footer_company_name || ''}">
                            </div>
                            <div class="form-group full-width">
                                <label>Alamat Lengkap</label>
                                <textarea name="footer_address" rows="4">${settings.footer_address || ''}</textarea>
                            </div>
                            <div class="form-group">
                                <label>Nomor Telepon Head Office</label>
                                <input type="text" name="footer_phone" value="${settings.footer_phone || ''}">
                            </div>
                            <div class="form-group">
                                <label>Jam Kerja / Operasional</label>
                                <input type="text" name="footer_working_hours" value="${settings.footer_working_hours || ''}">
                            </div>
                            <div class="form-group">
                                <label>URL LinkedIn</label>
                                <input type="url" name="footer_linkedin" value="${settings.footer_linkedin || ''}">
                            </div>
                            <div class="form-group">
                                <label>URL Facebook</label>
                                <input type="url" name="footer_facebook" value="${settings.footer_facebook || ''}">
                            </div>
                            <div class="form-group">
                                <label>URL Instagram</label>
                                <input type="url" name="footer_instagram" value="${settings.footer_instagram || ''}">
                            </div>
                            <div class="form-group">
                                <label>Teks Hak Cipta (Copyright)</label>
                                <input type="text" name="footer_copyright" value="${settings.footer_copyright || ''}">
                            </div>
                        </div>
                    </div>

                    <div class="section-card">
                        <h2><i class="fa-solid fa-envelope"></i> Bagian 4: Contact Us & Newsletter</h2>
                        <div class="form-grid">
                            <div class="form-group full-width">
                                <label>Judul Heading Form Contact Us</label>
                                <input type="text" name="contact_heading" value="${settings.contact_heading || 'Let us know what you need'}">
                            </div>
                            <div class="form-group full-width">
                                <label>Judul Newsletter</label>
                                <input type="text" name="news_newsletter_title" value="${settings.news_newsletter_title || ''}">
                            </div>
                            <div class="form-group full-width">
                                <label>Deskripsi / Subtitle Newsletter</label>
                                <textarea name="news_newsletter_text" rows="2">${settings.news_newsletter_text || ''}</textarea>
                            </div>
                        </div>
                    </div>

                    <div style="margin-bottom: 40px; text-align: right;">
                        <button type="submit" class="btn-save"><i class="fa-solid fa-floppy-disk"></i> Simpan Semua Perubahan</button>
                    </div>
                </form>

                <!-- SECTION KELOLA TRANSLATIONS (NON-HARDCODE) -->
                <div class="section-card">
                    <h2><i class="fa-solid fa-language"></i> Kelola Terjemahan Bahasa (En / Id)</h2>
                    <form action="/admin/cms/translations/update" method="POST">
                        <div class="table-responsive">
                            <table>
                                <thead>
                                    <tr>
                                        <th style="width: 25%;">Translation Key</th>
                                        <th style="width: 37.5%;">Bahasa Inggris (en_text)</th>
                                        <th style="width: 37.5%;">Bahasa Indonesia (id_text)</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${transRows}
                                </tbody>
                            </table>
                        </div>
                        <div style="margin-top: 20px; text-align: right;">
                            <button type="submit" class="btn-save"><i class="fa-solid fa-floppy-disk"></i> Simpan Terjemahan</button>
                        </div>
                    </form>
                </div>

                <div class="section-card">
                    <h2><i class="fa-solid fa-handshake"></i> Bagian 5: Kelola Our Partners (Marquee Slider)</h2>
                    <form action="/admin/partners/add" method="POST" enctype="multipart/form-data" style="margin-bottom: 25px;">
                        <div class="form-grid">
                            <div class="form-group">
                                <label>Nama Partner</label>
                                <input type="text" name="name" required placeholder="Contoh: Evyap">
                            </div>
                            <div class="form-group">
                                <label>URL Website Partner</label>
                                <input type="url" name="website_url" placeholder="https://example.com">
                            </div>
                            <div class="form-group full-width">
                                <label>File Logo Partner</label>
                                <input type="file" name="logo" accept="image/*" required>
                            </div>
                        </div>
                        <button type="submit" class="btn-save" style="background:#3182ce;"><i class="fa-solid fa-plus"></i> Tambah Partner Baru</button>
                    </form>

                    <div class="table-responsive">
                        <table>
                            <thead>
                                <tr><th style="width: 100px; text-align: center;">Logo</th><th>Nama Partner</th><th>URL Website</th><th style="width: 120px; text-align: center;">Aksi</th></tr>
                            </thead>
                            <tbody>${partnerRows.length > 0 ? partnerRows : '<tr><td colspan="4" style="text-align:center;">Belum ada partner.</td></tr>'}</tbody>
                        </table>
                    </div>
                </div>
            </div>

            <div id="editPartnerModal" class="modal-overlay">
                <div class="modal-box">
                    <h3 style="margin-top:0; color:#008744;">Edit Partner</h3>
                    <form id="editPartnerForm" method="POST" enctype="multipart/form-data">
                        <div class="form-group" style="margin-bottom:15px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px;">Nama Partner</label>
                            <input type="text" id="editPartnerName" name="name" required style="width:100%; padding:8px; box-sizing:border-box;">
                        </div>
                        <div class="form-group" style="margin-bottom:15px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px;">URL Website Partner</label>
                            <input type="url" id="editPartnerWebsite" name="website_url" style="width:100%; padding:8px; box-sizing:border-box;">
                        </div>
                        <div class="form-group" style="margin-bottom:15px;">
                            <label style="display:block; font-weight:600; margin-bottom:5px;">Ganti Logo Partner (Opsional)</label>
                            <div id="currentLogoPreview" style="margin-bottom:10px;"></div>
                            <input type="file" name="logo" accept="image/*">
                        </div>
                        <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:20px;">
                            <button type="button" onclick="closeEditPartnerModal()" style="padding:8px 16px; border:none; background:#cbd5e0; border-radius:4px; cursor:pointer;">Batal</button>
                            <button type="submit" style="padding:8px 16px; border:none; background:#008744; color:#fff; font-weight:600; border-radius:4px; cursor:pointer;">Simpan Perubahan</button>
                        </div>
                    </form>
                </div>
            </div>

            <script>
                function openEditPartnerModal(id, name, website, logoUrl) {
                    document.getElementById('editPartnerForm').action = '/admin/partners/edit/' + id;
                    document.getElementById('editPartnerName').value = name;
                    document.getElementById('editPartnerWebsite').value = website;
                    document.getElementById('currentLogoPreview').innerHTML = logoUrl ? '<img src="' + logoUrl + '" style="height:40px; border:1px solid #ddd; padding:3px; border-radius:4px;">' : '';
                    document.getElementById('editPartnerModal').classList.add('active');
                }
                function closeEditPartnerModal() {
                    document.getElementById('editPartnerModal').classList.remove('active');
                }
            </script>
        </body>
        </html>
    `);
});

// ==========================================
// API ENDPOINTS UNTUK TRANSLATIONS (NON-HARDCODE)
// ==========================================
app.get('/api/translations', (req, res) => {
    try {
        const result = db.exec(`SELECT translation_key, en_text, id_text FROM translations`);
        let translations = { en: {}, id: {} };
        
        if (result.length > 0) {
            result[0].values.forEach(row => {
                const [key, enText, idText] = row;
                translations.en[key] = enText;
                translations.id[key] = idText;
            });
        }
        res.json(translations);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/admin/cms/translations/update', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send('Unauthorized');
    
    try {
        const body = req.body;
        
        for (const [key, value] of Object.entries(body)) {
            if (key.startsWith('en_')) {
                const tKey = key.replace('en_', '');
                db.run(`UPDATE translations SET en_text = ? WHERE translation_key = ?`, [value, tKey]);
            } else if (key.startsWith('id_')) {
                const tKey = key.replace('id_', '');
                db.run(`UPDATE translations SET id_text = ? WHERE translation_key = ?`, [value, tKey]);
            }
        }
        
        saveDatabase();
        res.redirect('/admin/cms');
    } catch (err) {
        res.status(500).send('Gagal memperbarui terjemahan: ' + err.message);
    }
});

// Update Pengaturan CMS & Handle Multi-Upload File Gambar Lengkap
// Update Pengaturan CMS & Handle Multi-Upload File Gambar Lengkap
app.post('/admin/settings/update', (req, res) => {
    upload.fields([
        { name: 'logo_white_file', maxCount: 1 },
        { name: 'logo_dark_file', maxCount: 1 },
        { name: 'hero_bg_file_1', maxCount: 1 },
        { name: 'hero_bg_file_2', maxCount: 1 },
        { name: 'hero_bg_file_3', maxCount: 1 },
        { name: 'hero_bg_file_4', maxCount: 1 },
        { name: 'hero_bg_file_5', maxCount: 1 },
        { name: 'cta_bg_file', maxCount: 1 },
        { name: 'principals_bg_file', maxCount: 1 },
        { name: 'about_bg_file', maxCount: 1 },
        { name: 'about_tab_file_1', maxCount: 1 },
        { name: 'about_tab_file_2', maxCount: 1 },
        { name: 'about_tab_file_3', maxCount: 1 },
        { name: 'about_tab_file_4', maxCount: 1 },
        { name: 'contact_bg_file', maxCount: 1 },
        { name: 'news_bg_file', maxCount: 1 },
        { name: 'newsletter_bg_file', maxCount: 1 },
        { name: 'segment_personal_care_bg_file', maxCount: 1 },
        { name: 'segment_pharma_bg_file', maxCount: 1 },
        { name: 'segment_homecare_bg_file', maxCount: 1 },
        { name: 'segment_vet_bg_file', maxCount: 1 },
        { name: 'segment_processing_bg_file', maxCount: 1 },
        { name: 'app_pc_cosmetics_img_file', maxCount: 1 },
        { name: 'app_pc_skincare_img_file', maxCount: 1 },
        { name: 'app_pc_bath_img_file', maxCount: 1 },
        { name: 'app_pc_haircare_img_file', maxCount: 1 },
        { name: 'app_pharma_pharma_img_file', maxCount: 1 },
        { name: 'app_pharma_food_img_file', maxCount: 1 },
        { name: 'app_pharma_excipient_img_file', maxCount: 1 },
        { name: 'app_homecare_home_img_file', maxCount: 1 },
        { name: 'app_homecare_inst_img_file', maxCount: 1 },
        { name: 'app_homecare_ind_img_file', maxCount: 1 },
        { name: 'app_vet_aqua_img_file', maxCount: 1 },
        { name: 'app_vet_livestock_img_file', maxCount: 1 },
        { name: 'app_proc_metalworking_img_file', maxCount: 1 },
        { name: 'app_proc_battery_img_file', maxCount: 1 },
        { name: 'app_proc_textile_img_file', maxCount: 1 },
        { name: 'app_proc_water_img_file', maxCount: 1 },
        { name: 'app_proc_cropcare_img_file', maxCount: 1 }
    ])(req, res, function (err) {
        if (err) return res.status(500).send('Gagal mengunggah gambar: ' + err.message);
        if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send('Unauthorized');

        try {
            const formData = req.body || {};
            for (const [key, value] of Object.entries(formData)) {
                if (value !== undefined && value !== null) {
                    db.run(`INSERT OR REPLACE INTO site_settings (key, value) VALUES (?, ?)`, [key, value]);
                }
            }

            const fileKeys = [
                { field: 'logo_white_file', dbKey: 'logo_white' },
                { field: 'logo_dark_file', dbKey: 'logo_dark' },
                { field: 'hero_bg_file_1', dbKey: 'hero_bg_image_1' },
                { field: 'hero_bg_file_2', dbKey: 'hero_bg_image_2' },
                { field: 'hero_bg_file_3', dbKey: 'hero_bg_image_3' },
                { field: 'hero_bg_file_4', dbKey: 'hero_bg_image_4' },
                { field: 'hero_bg_file_5', dbKey: 'hero_bg_image_5' },
                { field: 'cta_bg_file', dbKey: 'cta_bg_image' },
                { field: 'principals_bg_file', dbKey: 'principals_bg_image' },
                { field: 'contact_bg_file', dbKey: 'contact_bg_image' },
                { field: 'about_bg_file', dbKey: 'about_bg_image' },
                { field: 'about_tab_file_1', dbKey: 'about_tab_img_1' },
                { field: 'about_tab_file_2', dbKey: 'about_tab_img_2' },
                { field: 'about_tab_file_3', dbKey: 'about_tab_img_3' },
                { field: 'about_tab_file_4', dbKey: 'about_tab_img_4' },
                { field: 'news_bg_file', dbKey: 'news_bg_image' },
                { field: 'newsletter_bg_file', dbKey: 'newsletter_bg_image' },
                { field: 'segment_personal_care_bg_file', dbKey: 'segment_personal_care_bg' },
                { field: 'segment_pharma_bg_file', dbKey: 'segment_pharma_bg' },
                { field: 'segment_homecare_bg_file', dbKey: 'segment_homecare_bg' },
                { field: 'segment_vet_bg_file', dbKey: 'segment_vet_bg' },
                { field: 'segment_processing_bg_file', dbKey: 'segment_processing_bg' },
                { field: 'app_pc_cosmetics_img_file', dbKey: 'app_pc_cosmetics_img' },
                { field: 'app_pc_skincare_img_file', dbKey: 'app_pc_skincare_img' },
                { field: 'app_pc_bath_img_file', dbKey: 'app_pc_bath_img' },
                { field: 'app_pc_haircare_img_file', dbKey: 'app_pc_haircare_img' },
                { field: 'app_pharma_pharma_img_file', dbKey: 'app_pharma_pharma_img' },
                { field: 'app_pharma_food_img_file', dbKey: 'app_pharma_food_img' },
                { field: 'app_pharma_excipient_img_file', dbKey: 'app_pharma_excipient_img' },
                { field: 'app_homecare_home_img_file', dbKey: 'app_homecare_home_img' },
                { field: 'app_homecare_inst_img_file', dbKey: 'app_homecare_inst_img' },
                { field: 'app_homecare_ind_img_file', dbKey: 'app_homecare_ind_img' },
                { field: 'app_vet_aqua_img_file', dbKey: 'app_vet_aqua_img' },
                { field: 'app_vet_livestock_img_file', dbKey: 'app_vet_livestock_img' },
                { field: 'app_proc_metalworking_img_file', dbKey: 'app_proc_metalworking_img' },
                { field: 'app_proc_battery_img_file', dbKey: 'app_proc_battery_img' },
                { field: 'app_proc_textile_img_file', dbKey: 'app_proc_textile_img' },
                { field: 'app_proc_water_img_file', dbKey: 'app_proc_water_img' },
                { field: 'app_proc_cropcare_img_file', dbKey: 'app_proc_cropcare_img' }
            ];

            if (req.files) {
                fileKeys.forEach(item => {
                    if (req.files[item.field] && req.files[item.field][0]) {
                        const imgPath = '/uploads/' + req.files[item.field][0].filename;
                        db.run(`INSERT OR REPLACE INTO site_settings (key, value) VALUES (?, ?)`, [item.dbKey, imgPath]);
                    }
                });
            }

            // PANGGIL saveDatabase() AGAR TERSIMPAN KE FILE FISIK TRIAARTH.SQLITE
            saveDatabase();
            
            res.redirect('/admin/cms');
        } catch (dbErr) {
            res.status(500).send('Gagal memperbarui pengaturan: ' + dbErr.message);
        }
    });
});

app.post('/admin/partners/add', upload.single('logo'), (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send('Unauthorized');
    const { name, website_url } = req.body;
    const logoUrl = req.file ? `/uploads/${req.file.filename}` : '';

    try {
        db.run(`INSERT INTO partners (name, logo_url, website_url) VALUES (?, ?, ?)`, [name, logoUrl, website_url]);
        saveDatabase();
        res.redirect('/admin/cms');
    } catch (err) {
        res.status(500).send('Gagal menambah partner: ' + err.message);
    }
});

app.post('/admin/partners/edit/:id', upload.single('logo'), (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send('Unauthorized');

    const { name, website_url } = req.body;
    const partnerId = req.params.id;

    try {
        if (req.file) {
            const logoUrl = `/uploads/${req.file.filename}`;
            db.run(`UPDATE partners SET name = ?, website_url = ?, logo_url = ? WHERE id = ?`, [name, website_url, logoUrl, partnerId]);
        } else {
            db.run(`UPDATE partners SET name = ?, website_url = ? WHERE id = ?`, [name, website_url, partnerId]);
        }
        saveDatabase();
        res.redirect('/admin/cms');
    } catch (err) {
        res.status(500).send('Gagal memperbarui partner: ' + err.message);
    }
});

app.post('/admin/partners/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send('Unauthorized');
    try {
        db.run(`DELETE FROM partners WHERE id = ?`, [req.params.id]);
        saveDatabase();
        res.redirect('/admin/cms');
    } catch (err) {
        res.status(500).send('Gagal menghapus partner: ' + err.message);
    }
});

// API untuk mengambil SEMUA produk lintas segmen
app.get('/api/all-products', async (req, res) => {
    try {
        const result = await db.execute(`SELECT * FROM products`);
        const products = result.rows.map(row => {
            if (!req.session.user) row.formula_pdf = null;
            return row;
        });
        res.json(products);
    } catch (err) {
        res.status(500).json({ error: 'Gagal mengambil data produk' });
    }
});

// API untuk mengambil opsi filter global (Principal, Segment, Function, Application)
app.get('/api/all-filter-options', (req, res) => {
    try {
        const result = db.exec(`SELECT DISTINCT principal, segment, function, application FROM products`);
        let rows = [];
        if (result.length > 0) {
            const columns = result[0].columns;
            rows = result[0].values.map(row => {
                let obj = {};
                columns.forEach((col, i) => obj[col] = row[i]);
                return obj;
            });
        }
        
        const principals = [...new Set(rows.map(r => r.principal).filter(Boolean))].sort();
        const segments = [...new Set(rows.map(r => r.segment).filter(Boolean))].sort();
        const functions = [...new Set(rows.map(r => r.function).filter(Boolean))].sort();
        const applications = [...new Set(rows.map(r => r.application).filter(Boolean))].sort();

        res.json({ principals, segments, functions, applications });
    } catch (err) {
        console.error(err.message);
        res.status(500).json({ error: 'Gagal mengambil opsi filter' });
    }
});

// ==========================================
// API PUBLIK
// ==========================================
app.get('/api/site-settings', (req, res) => {
    try {
        const result = db.exec(`SELECT * FROM site_settings`);
        let settings = {};
        if (result.length > 0) {
            result[0].values.forEach(row => { settings[row[0]] = row[1]; });
        }
        res.json(settings);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/partners', (req, res) => {
    try {
        const result = db.exec(`SELECT * FROM partners ORDER BY id DESC`);
        let partners = [];
        if (result.length > 0) {
            const cols = result[0].columns;
            partners = result[0].values.map(row => {
                let obj = {}; cols.forEach((c, i) => obj[c] = row[i]); return obj;
            });
        }
        res.json(partners);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/news', (req, res) => {
    try {
        const page = parseInt(req.query.page) || 1;
        const limit = 6; // 6 berita per halaman
        const offset = (page - 1) * limit;

        // Hitung total berita keseluruhan
        const countStmt = db.prepare(`SELECT COUNT(*) as total FROM news`);
        let totalNews = 0;
        if (countStmt.step()) {
            totalNews = countStmt.getAsObject().total;
        }
        countStmt.free();

        // Ambil data berita sesuai halaman (LIMIT & OFFSET)
        const stmt = db.prepare(`
            SELECT news.*, users.name as author_name 
            FROM news 
            JOIN users ON news.author_id = users.id 
            ORDER BY created_at DESC 
            LIMIT ? OFFSET ?
        `);
        stmt.bind([limit, offset]);

        let newsList = [];
        while (stmt.step()) {
            newsList.push(stmt.getAsObject());
        }
        stmt.free();

        res.json({
            news: newsList,
            currentPage: page,
            totalPages: Math.ceil(totalNews / limit)
        });
    } catch (err) {
        res.status(500).json({ error: 'Gagal memuat berita' });
    }
});

app.get('/api/news/:slug', (req, res) => {
    try {
        const stmt = db.prepare(`SELECT news.*, users.name as author_name FROM news JOIN users ON news.author_id = users.id WHERE news.slug = ?`);
        stmt.bind([req.params.slug]);

        if (stmt.step()) {
            const article = stmt.getAsObject();
            stmt.free();
            res.json(article);
        } else {
            stmt.free();
            res.status(404).json({ error: 'Berita tidak ditemukan' });
        }
    } catch (err) {
        res.status(500).json({ error: 'Kesalahan server' });
    }
});

app.get('/api/check-session', (req, res) => {
    if (req.session.user) {
        res.json({ loggedIn: true, role: req.session.user.role, name: req.session.user.name });
    } else {
        res.json({ loggedIn: false });
    }
});

app.get('/products/:slug', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'product-detail.html'));
});

app.get('/api/products/:slug', (req, res) => {
    try {
        const stmt = db.prepare(`SELECT * FROM products WHERE slug = ?`);
        stmt.bind([req.params.slug]);

        if (stmt.step()) {
            const product = stmt.getAsObject();
            stmt.free();

            if (!req.session.user) {
                product.formula_pdf = null;
            }
            res.json(product);
        } else {
            stmt.free();
            res.status(404).json({ error: 'Produk tidak ditemukan' });
        }
    } catch (err) {
        res.status(500).json({ error: 'Kesalahan server' });
    }
});

app.get('/logout', (req, res) => {
    req.session.destroy(() => res.redirect('/signin'));
});

app.post('/admin/products/upload-by-segment', upload.single('product_file'), (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    if (!req.file) return res.status(400).send('File Excel/CSV tidak ditemukan.');

    const defaultSegment = req.body.segment || 'Personal Care, Cosmetic & Fragrance';
    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        let insertedCount = 0;
        sheetData.forEach(row => {
            const segment = row.Segment || row.segment || defaultSegment;
            const principal = row.Principal || row.principal || '-';
            const productName = row['Product Name'] || row.product_name || '';
            const inciName = row['INCI Name'] || row.inci_name || '';
            const description = row.Description || row.description || '';
            const func = row.Function || row.function || '';
            const application = row.Application || row.application || '';
            const formulaPdf = row['Formula PDF'] || row.formula_pdf || '';

            if (productName) {
                const checkStmt = db.prepare(`SELECT COUNT(*) as count FROM products WHERE product_name = ? AND segment = ?`);
                checkStmt.bind([productName, segment]);
                let exists = false;
                if (checkStmt.step()) {
                    if (checkStmt.getAsObject().count > 0) {
                        exists = true;
                    }
                }
                checkStmt.free();

                if (!exists) {
                    const slug = productName.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') + '-' + Math.floor(Math.random() * 10000);
                    db.run(`INSERT INTO products (segment, principal, product_name, slug, inci_name, description, function, application, formula_pdf) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                        [segment, principal, productName, slug, inciName, description, func, application, formulaPdf]
                    );
                    insertedCount++;
                }
            }
        });

        saveDatabase();
        fs.unlinkSync(req.file.path);
        res.send(renderMessagePage('Berhasil!', `Data produk dari file Excel berhasil di-upload (${insertedCount} produk baru ditambahkan, duplikat dilewati).`, 'success', '/admin/products/dashboard', 'Kembali ke Kelola Produk'));
    } catch (err) {
        res.status(500).send('Gagal memproses file Excel/CSV: ' + err.message);
    }
});

app.post('/admin/upload-product-pdf', upload.single('formula_pdf'), (req, res) => {
    if (!isEditorOrAdmin(req)) return res.status(403).send('Unauthorized');
    if (!req.file) return res.status(400).send('File PDF tidak ditemukan.');

    const { segment, product_name } = req.body;
    const pdfUrl = `/uploads/${req.file.filename}`;
    try {
        db.run(`INSERT INTO product_formulas (product_name, segment, pdf_url) VALUES (?, ?, ?)`, [product_name, segment, pdfUrl]);
        db.run(`UPDATE products SET formula_pdf = ? WHERE product_name = ? AND segment = ?`, [pdfUrl, product_name, segment]);
        
        saveDatabase();
        res.send(renderMessagePage('Berhasil!', `Dokumen PDF formula untuk produk <b>${product_name}</b> (${segment}) berhasil diunggah dan dihubungkan.`, 'success', '/admin/products/dashboard', 'Kembali ke Kelola Produk'));
    } catch (err) {
        res.status(500).send('Gagal menyimpan PDF: ' + err.message);
    }
});

app.get('/api/products-by-segment', (req, res) => {
    const { segment } = req.query;
    try {
        let query = `SELECT * FROM products`;
        let params = [];
        if (segment) {
            query += ` WHERE segment = ?`;
            params.push(segment);
        }
        const stmt = db.prepare(query);
        if (params.length > 0) {
            stmt.bind(params);
        }
        
        let products = [];
        while (stmt.step()) {
            let prod = stmt.getAsObject();
            if (!req.session.user) {
                prod.formula_pdf = null;
            }
            products.push(prod);
        }
        stmt.free();
        res.json(products);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/filter-options', (req, res) => {
    const { segment } = req.query;
    try {
        let queryParams = [];
        let whereClause = '';
        if (segment) {
            whereClause = `WHERE segment = ?`;
            queryParams.push(segment);
        }

        const getValues = (sql, params) => {
            const stmt = db.prepare(sql);
            if (params.length > 0) stmt.bind(params);
            let results = [];
            while (stmt.step()) {
                results.push(stmt.getAsObject());
            }
            stmt.free();
            return results;
        };

        const principalsRes = getValues(`SELECT DISTINCT principal FROM products ${whereClause}`, queryParams);
        const functionsRes = getValues(`SELECT DISTINCT function FROM products ${whereClause}`, queryParams);
        const applicationsRes = getValues(`SELECT DISTINCT application FROM products ${whereClause}`, queryParams);

        // Helper function dengan parameter allowDash (default true atau diatur true agar '-' lolos)
        const extractValues = (resultList, colName, allowDash = true) => {
            return [...new Set(resultList.map(row => row[colName]).filter(val => {
                if (val === null || val === undefined) return false;
                if (!allowDash && val === '-') return false;
                return true;
            }))].sort();
        };

        res.json({
            principals: extractValues(principalsRes, 'principal', true), // Izinkan '-' untuk principals
            functions: extractValues(functionsRes, 'function', true),       // Izinkan '-' jika diperlukan
            applications: extractValues(applicationsRes, 'application', true) // Izinkan '-' jika diperlukan
        });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ==========================================
// JALANKAN SERVER
// ==========================================
app.listen(PORT, () => {
    console.log(`Server aktif di http://localhost:${PORT}`);
});