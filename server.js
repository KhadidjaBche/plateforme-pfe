require('dotenv').config();
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const cors = require('cors');
const path = require('path');
const multer = require('multer'); // Gestionnaire de téléchargement de fichiers
const xlsx = require('xlsx');     // Analyseur de fichiers Excel
const fs = require('fs');         // Gestionnaire de fichiers système
const bcrypt = require('bcryptjs'); // Bibliothèque pour crypter les mots de passe
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;
let schemaReady = false;

// Sessions en mémoire : adaptées au fonctionnement local de cette version.
// Pour un déploiement multi-instance, utiliser un stockage de sessions persistant.
const activeSessions = new Map();

// Normalisation commune des valeurs textuelles : espaces périphériques et casse.
// Les valeurs affichées/conservées en base gardent leur écriture d'origine.
function normalizeText(value) {
    return String(value == null ? '' : value).normalize('NFKC').trim().toLocaleLowerCase('fr-FR');
}
function normalizeMatricule(value) {
    return normalizeText(value);
}
function sameText(a, b) {
    return normalizeText(a) === normalizeText(b);
}
function canonicalLevel(value) {
    const v = normalizeText(value).replace(/\s+/g, ' ');
    const levels = {
        'l3': 'L3',
        'licence 3': 'L3',
        'licence3': 'L3',
        '3eme annee licence': 'L3',
        '3ème année licence': 'L3',
        '3e annee licence': 'L3',
        '3e année licence': 'L3',
        '4ing': '4ING',
        '4 ing': '4ING',
        'master 2': 'Master 2',
        'm2': 'Master 2',
        '5ing': '5ING',
        '5 ing': '5ING',
        'm2 + 5ing': 'M2 + 5ING',
        'master 2 + 5ing': 'M2 + 5ING'
    };
    return levels[v] || String(value == null ? '' : value).normalize('NFKC').trim();
}

function requireAuth(req, res, next) {
    const authorization = req.headers.authorization || '';
    const token = authorization.startsWith('Bearer ') ? authorization.slice(7) : null;
    const sessionUser = token ? activeSessions.get(token) : null;
    if (!sessionUser) return res.status(401).json({ error: 'Session absente ou expirée. Veuillez vous reconnecter.' });
    req.user = sessionUser;
    next();
}
function requireAdmin(req, res, next) {
    if (!sameText(req.user.role, 'admin')) return res.status(403).json({ error: 'Accès réservé à l’administrateur.' });
    next();
}
function requireTeacher(req, res, next) {
    if (!['teacher', 'validator'].includes(normalizeText(req.user.role)) || !req.user.teacher_id) {
        return res.status(403).json({ error: 'Accès réservé aux enseignants disposant d’un profil importé.' });
    }
    const requestedId = req.params.id || req.body.enseignant_id;
    if (requestedId && Number(requestedId) !== Number(req.user.teacher_id)) {
        return res.status(403).json({ error: 'Vous ne pouvez modifier que votre propre profil enseignant.' });
    }
    next();
}
function requireValidator(req, res, next) {
    if (!req.user.teacher_id || Number(req.user.is_validator) !== 1) {
        return res.status(403).json({ error: 'Accès réservé aux enseignants désignés comme validateurs.' });
    }
    next();
}


// ==========================================
// CONFIGURATION DES MIDDLEWARES
// ==========================================
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req,res,next) => { const indexPath=path.join(__dirname,'index.html'); fs.existsSync(indexPath) ? res.sendFile(indexPath) : next(); });
app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use('/api', (req, res, next) => schemaReady ? next() : res.status(503).json({error:'La base de données est en cours d’initialisation. Réessayez dans quelques secondes.'}));

// Configuration de Multer pour stocker temporairement les fichiers Excel uploadés
const upload = multer({
    dest: path.join(__dirname, 'uploads'),
    limits: { fileSize: 10 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (['.xlsx', '.xls'].includes(ext)) cb(null, true);
        else cb(new Error('Seuls les fichiers Excel .xlsx et .xls sont autorisés.'));
    }
});

// ==========================================
// CONFIGURATION DE MULTER POUR LES SUJETS (PDF/DOC)
// ==========================================
const sujetsDir = path.join(__dirname, 'uploads', 'sujets');
if (!fs.existsSync(sujetsDir)) {
    fs.mkdirSync(sujetsDir, { recursive: true });
}

const storageSujet = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, sujetsDir);
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, 'sujet-' + uniqueSuffix + path.extname(file.originalname));
    }
});

const uploadSujet = multer({
    storage: storageSujet,
    limits: { fileSize: 5 * 1024 * 1024 }, // Limite à 5 Mo
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['.pdf', '.doc', '.docx'];
        const ext = path.extname(file.originalname).toLowerCase();
        const allowedMime = {
            '.pdf': ['application/pdf', 'application/octet-stream'],
            '.doc': ['application/msword', 'application/octet-stream'],
            '.docx': ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/octet-stream']
        };
        if (allowedTypes.includes(ext) && allowedMime[ext].includes(file.mimetype)) cb(null, true);
        else cb(new Error('Format de fichier non autorisé. Utilisez un vrai fichier PDF, DOC ou DOCX.'));
    }
});

// ==========================================
// INITIALISATION DE LA BASE DE DONNÉES
// ==========================================
const dbPath = path.resolve(__dirname, 'database.sqlite');
const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error('Erreur de connexion à la base de données :', err.message);
    } else {
        console.log('Connecté à la base de données SQLite avec succès.');
    }
});

// ==========================================
// CRÉATION DES TABLES (Cahier des charges)
// ==========================================
db.serialize(() => {
    db.run('PRAGMA foreign_keys = ON');

    // 1. Table des Utilisateurs globaux (Admin, Étudiants, Enseignants, Validateurs)
    db.run(`CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('admin', 'student', 'teacher', 'validator')),
        status TEXT DEFAULT 'pending' CHECK(status IN ('pending', 'approved', 'rejected', 'disabled')),
        email TEXT UNIQUE NOT NULL,
        nom TEXT NOT NULL,
        prenom TEXT NOT NULL,
        matricule TEXT UNIQUE,
        niveau TEXT,
        specialite TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);

    // 2. Table des Spécialités (L3, 4ING, M2, 5ING)
    db.run(`CREATE TABLE IF NOT EXISTS specialites (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        niveau TEXT NOT NULL,
        nom_specialite TEXT NOT NULL
    )`);

    // 3. Table des Étudiants (Importés via Excel)
    db.run(`CREATE TABLE IF NOT EXISTS etudiants_import (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        annee_universitaire TEXT NOT NULL,
        niveau TEXT NOT NULL,
        specialite TEXT NOT NULL,
        numero_ordre INTEGER,
        numero_inscription TEXT,
        annee_bac TEXT,
        matricule TEXT UNIQUE NOT NULL,
        nom TEXT NOT NULL,
        prenom TEXT NOT NULL,
        section TEXT,
        groupe TEXT
    )`);

    // 4. Table des Enseignants (Importés via Excel)
    db.run(`CREATE TABLE IF NOT EXISTS enseignants_import (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        annee_universitaire TEXT NOT NULL,
        nom TEXT NOT NULL,
        prenom TEXT NOT NULL,
        email TEXT UNIQUE NOT NULL,
        is_validator INTEGER DEFAULT 0,
        max_groupes_l3 INTEGER DEFAULT 0,
        max_groupes_4ing INTEGER DEFAULT 0,
        max_groupes_m2_5ing INTEGER DEFAULT 2
    )`);

    // 5. Table des Groupes de projets
    db.run(`CREATE TABLE IF NOT EXISTS groupes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        annee_universitaire TEXT NOT NULL,
        niveau TEXT NOT NULL,
        specialite TEXT NOT NULL,
        type_groupe TEXT NOT NULL,
        sujet_titre TEXT,
        sujet_description_path TEXT,
        statut_sujet TEXT DEFAULT 'en_attente' CHECK(statut_sujet IN ('en_attente', 'valide', 'invalide')),
        enseignant_id INTEGER,
        FOREIGN KEY(enseignant_id) REFERENCES enseignants_import(id)
    )`);

    // 6. Table des Membres du groupe (Liaison Groupe <-> Étudiants)
    db.run(`CREATE TABLE IF NOT EXISTS groupe_membres (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        groupe_id INTEGER NOT NULL,
        etudiant_id INTEGER NOT NULL UNIQUE,
        role_dans_groupe TEXT DEFAULT 'membre' CHECK(role_dans_groupe IN ('chef_groupe', 'membre')),
        FOREIGN KEY(groupe_id) REFERENCES groupes(id) ON DELETE CASCADE,
        FOREIGN KEY(etudiant_id) REFERENCES etudiants_import(id) ON DELETE CASCADE
    )`);

    // 7. Table des Sujets Internes (Proposés par les enseignants)
    db.run(`CREATE TABLE IF NOT EXISTS sujets_internes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        annee_universitaire TEXT NOT NULL,
        enseignant_id INTEGER NOT NULL,
        niveau TEXT NOT NULL,
        specialite TEXT,
        titre TEXT NOT NULL,
        description TEXT,
        capacite_max INTEGER DEFAULT 1,
        FOREIGN KEY(enseignant_id) REFERENCES enseignants_import(id) ON DELETE CASCADE
    )`);

    console.log('Toutes les tables de la base de données sont prêtes.');
});

// Protection des API selon le rôle connecté.
app.use('/api/admin', requireAuth, requireAdmin);
app.use('/api/enseignants', requireAuth, requireTeacher);
app.use('/api/validateur', requireAuth, requireValidator);

// ==========================================
// ROUTES API - GESTION DES ÉTUDIANTS (EXCEL)
// ==========================================
app.post('/api/admin/import-etudiants', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'Veuillez fournir un fichier Excel valide.' });
    }

    const { annee_universitaire, niveau, specialite } = req.body;
    const anneeImport = String(annee_universitaire || '').trim();
    const niveauImport = canonicalLevel(niveau);
    const specialiteImport = String(specialite || '').trim();
    if (!anneeImport || !niveauImport || !specialiteImport) {
        return res.status(400).json({ error: 'Veuillez préciser l\'année universitaire, le niveau et la spécialité.' });
    }

    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        if (sheetData.length === 0) {
            return res.status(400).json({ error: 'Le fichier Excel est vide.' });
        }

        let insertedCount = 0;
        let errorsCount = 0;

        db.serialize(() => {
            db.run('BEGIN TRANSACTION');

            const query = `INSERT OR IGNORE INTO etudiants_import 
                (annee_universitaire, niveau, specialite, numero_ordre, numero_inscription, annee_bac, matricule, nom, prenom, section, groupe) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

            sheetData.forEach((row) => {
                const numero_ordre = row['N°'] || row['Numero'] || null;
                const numero_inscription = row['Numero Inscription'] || row['Numero_Inscription'] || '';
                const annee_bac = row['Annee Bac'] || row['Annee_Bac'] || '';
                const matricule = row['Matricule'] == null ? '' : String(row['Matricule']).trim();
                const nom = row['Nom'] == null ? '' : String(row['Nom']).trim();
                const prenom = String(row['Prénom'] || row['Prenom'] || '').trim();
                const section = row['Section'] || '';
                const groupe = row['Groupe'] || '';

                if (matricule && nom && prenom) {
                    db.run(query, [
                        anneeImport,
                        niveauImport,
                        specialiteImport,
                        numero_ordre,
                        String(numero_inscription),
                        String(annee_bac),
                        String(matricule),
                        nom,
                        prenom,
                        String(section),
                        String(groupe)
                    ], function(err) {
                        if (err || this.changes === 0) errorsCount++;
                        else insertedCount++;
                    });
                } else {
                    errorsCount++;
                }
            });

            db.run('COMMIT', (err) => {
                if (err) {
                    return res.status(500).json({ error: 'Erreur lors de la validation de la transaction.' });
                }
                res.json({
                    message: 'Importation des étudiants terminée avec succès.',
                    details: {
                        inseres: insertedCount,
                        ignores_ou_erreurs: errorsCount
                    }
                });
            });
        });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Erreur lors du traitement du fichier Excel.' });
    }
});

// ==========================================
// ROUTES API - GESTION DES ENSEIGNANTS (EXCEL)
// ==========================================
app.post('/api/admin/import-enseignants', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'Veuillez fournir un fichier Excel valide.' });
    }

    const annee_universitaire = String(req.body.annee_universitaire || '').trim();
    if (!annee_universitaire) {
        return res.status(400).json({ error: 'Veuillez préciser l\'année universitaire.' });
    }

    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        if (sheetData.length === 0) {
            return res.status(400).json({ error: 'Le fichier Excel est vide.' });
        }

        let insertedCount = 0;
        let errorsCount = 0;

        db.serialize(() => {
            db.run('BEGIN TRANSACTION');

            const query = `INSERT OR IGNORE INTO enseignants_import (annee_universitaire, nom, prenom, email, max_groupes_m2_5ing) VALUES (?, ?, ?, ?, 2)`;

            sheetData.forEach((row) => {
                const nom = row['Nom'] == null ? '' : String(row['Nom']).trim();
                const prenom = String(row['Prénom'] || row['Prenom'] || '').trim();
                const email = String(row['Email'] || '').trim().toLowerCase();

                if (nom && prenom && email) {
                    db.run(query, [annee_universitaire, nom, prenom, email], function(err) {
                        if (err || this.changes === 0) errorsCount++;
                        else insertedCount++;
                    });
                } else {
                    errorsCount++;
                }
            });

            db.run('COMMIT', (err) => {
                if (err) {
                    return res.status(500).json({ error: 'Erreur lors de la transaction.' });
                }
                res.json({
                    message: 'Importation des enseignants terminée.',
                    details: {
                        inseres: insertedCount,
                        ignores_ou_erreurs: errorsCount
                    }
                });
            });
        });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Erreur lors du traitement du fichier Excel des enseignants.' });
    }
});

// ==========================================
// ROUTES API - AUTHENTIFICATION ET GESTION DES COMPTES
// ==========================================
app.post('/api/auth/register', async (req, res) => {
    const { username, email, nom, prenom, matricule, niveau, specialite, role } = req.body;
    const cleanUsername = String(username || '').trim();
    const cleanMatricule = String(matricule || '').trim();
    const cleanNiveau = niveau ? canonicalLevel(niveau) : '';
    const cleanSpecialite = String(specialite || '').trim();
    const cleanNom = String(nom || '').trim();
    const cleanPrenom = String(prenom || '').trim();
    const cleanRole = normalizeText(role);
    const normalizedEmail = String(email || '').trim().toLowerCase();
    if (!cleanUsername || !normalizedEmail || !cleanNom || !cleanPrenom || !cleanRole) return res.status(400).json({ error: 'Veuillez remplir tous les champs obligatoires.' });
    if (!/^\S+@\S+\.\S+$/.test(normalizedEmail)) return res.status(400).json({ error: 'Adresse e-mail invalide.' });
    if (!['student', 'teacher'].includes(cleanRole)) return res.status(400).json({ error: 'Rôle invalide.' });
    let nodemailer;
    try { nodemailer = require('nodemailer'); } catch (_) { return res.status(503).json({error:'Le module nodemailer manque. Installez-le avec npm install nodemailer pour activer la vérification e-mail.'}); }
    if (!process.env.SMTP_HOST || !process.env.SMTP_USER || !process.env.SMTP_PASS) return res.status(503).json({error:'Le serveur e-mail n’est pas configuré (SMTP_HOST, SMTP_USER, SMTP_PASS). Aucun compte n’a été créé.'});
    try {
      const proceed = async () => {
        const temporaryPassword = crypto.randomBytes(9).toString('base64url');
        const verificationCode = String(crypto.randomInt(100000, 999999));
        const passwordHash = await bcrypt.hash(temporaryPassword, 12);
        const codeHash = crypto.createHash('sha256').update(verificationCode).digest('hex');
        const expires = new Date(Date.now()+24*60*60*1000).toISOString();
        const tempExpires = new Date(Date.now()+72*60*60*1000).toISOString();
        const transporter = nodemailer.createTransport({host:process.env.SMTP_HOST,port:Number(process.env.SMTP_PORT||587),secure:String(process.env.SMTP_SECURE||'false')==='true',auth:{user:process.env.SMTP_USER,pass:process.env.SMTP_PASS}});
        db.run(`INSERT INTO users(username,password,role,status,email,nom,prenom,matricule,niveau,specialite,email_verification_code_hash,email_verification_expires,temporary_password_expires,must_change_password,email_verified)
                VALUES(?,?,?,'pending',?,?,?,?,?,?,?,?,?,1,0)`,
          [cleanUsername,passwordHash,cleanRole,normalizedEmail,cleanNom,cleanPrenom,cleanMatricule||null,cleanNiveau||null,cleanSpecialite||null,codeHash,expires,tempExpires], async function(err) {
            if (err) return res.status(409).json({error:'Nom d’utilisateur, adresse e-mail ou matricule déjà utilisé.'});
            try {
              await transporter.sendMail({from:process.env.SMTP_FROM||process.env.SMTP_USER,to:normalizedEmail,subject:'Inscription à la plateforme PFE',text:`Bonjour ${prenom} ${nom},\n\nVotre nom d’utilisateur : ${username}\nMot de passe temporaire : ${temporaryPassword}\nCode de vérification e-mail : ${verificationCode}\n\nLe code expire dans 24 heures et le mot de passe temporaire dans 72 heures. Après vérification de l’adresse et validation administrative, vous devrez choisir un nouveau mot de passe.`});
              return res.status(201).json({message:'Un e-mail contenant vos identifiants temporaires et le code de vérification a été envoyé. Vérifiez votre boîte de réception.'});
            } catch(mailErr) {
              db.run('DELETE FROM users WHERE id=?',[this.lastID]);
              console.error('SMTP:',mailErr.message);
              return res.status(502).json({error:'Envoi e-mail impossible. Vérifiez la configuration SMTP et réessayez.'});
            }
          });
      };
      if (cleanRole === 'student') {
        if (!cleanMatricule) return res.status(400).json({error:'Le matricule est obligatoire.'});
        db.get('SELECT * FROM etudiants_import WHERE LOWER(TRIM(matricule))=LOWER(TRIM(?))',[cleanMatricule],(e,row)=>{if(e||!row)return res.status(400).json({error:'Matricule absent de la liste officielle.'}); if(cleanNiveau&&!sameText(cleanNiveau,row.niveau))return res.status(400).json({error:'Le niveau ne correspond pas au matricule importé.'}); if(cleanSpecialite&&!sameText(cleanSpecialite,row.specialite))return res.status(400).json({error:'La spécialité ne correspond pas au matricule importé.'}); proceed();});
      } else db.get('SELECT * FROM enseignants_import WHERE LOWER(TRIM(email))=LOWER(TRIM(?))',[normalizedEmail],(e,row)=>{if(e||!row)return res.status(400).json({error:'Cet e-mail ne figure pas dans la liste officielle des enseignants.'});proceed();});
    } catch (error) { console.error(error); res.status(500).json({error:'Erreur serveur lors de l’inscription.'}); }
});

app.post('/api/auth/verify-email', (req,res)=>{
  const {email,code}=req.body; if(!email||!code)return res.status(400).json({error:'E-mail et code obligatoires.'});
  const hash=crypto.createHash('sha256').update(String(code)).digest('hex');
  db.run(`UPDATE users SET email_verified=1,email_verification_code_hash=NULL,email_verification_expires=NULL WHERE LOWER(TRIM(email))=LOWER(TRIM(?)) AND email_verification_code_hash=? AND datetime(email_verification_expires)>CURRENT_TIMESTAMP`,[String(email).trim().toLowerCase(),hash],function(e){if(e)return res.status(500).json({error:'Erreur.'});if(!this.changes)return res.status(400).json({error:'Code incorrect ou expiré.'});res.json({message:'Adresse e-mail vérifiée. Votre compte attend la validation administrative.'});});
});
app.post('/api/auth/change-temporary-password', async (req,res)=>{
  const {username,temporary_password,new_password}=req.body;
  if(!username||!temporary_password||!new_password||String(new_password).length<10)return res.status(400).json({error:'Mot de passe requis (au moins 10 caractères).'});
  db.get('SELECT * FROM users WHERE LOWER(TRIM(username))=LOWER(TRIM(?))',[String(username).trim()],async(e,u)=>{
    if(e||!u)return res.status(400).json({error:'Identifiants invalides.'});
    if(!u.email_verified)return res.status(403).json({error:'Vérifiez d’abord votre adresse e-mail.'});
    if(!u.temporary_password_expires||new Date(u.temporary_password_expires)<new Date())return res.status(403).json({error:'Mot de passe temporaire expiré. Contactez l’administrateur.'});
    if(!await bcrypt.compare(temporary_password,u.password))return res.status(401).json({error:'Mot de passe temporaire incorrect.'});
    const hash=await bcrypt.hash(new_password,12);
    db.run('UPDATE users SET password=?,must_change_password=0,temporary_password_expires=NULL WHERE id=?',[hash,u.id],function(err){err?res.status(500).json({error:'Mise à jour impossible.'}):res.json({message:'Mot de passe définitif enregistré. Vous pouvez vous connecter.'});});
  });
});

app.post('/api/admin/users/:id/validate', (req, res) => {
    const userId = req.params.id;
    const { action } = req.body;
    if (action === 'approve') {
        return db.get('SELECT email_verified FROM users WHERE id=?', [userId], (checkErr, row) => {
            if (checkErr) return res.status(500).json({error:'Vérification du compte impossible.'});
            if (!row) return res.status(404).json({error:'Utilisateur non trouvé.'});
            if (Number(row.email_verified) !== 1) return res.status(400).json({error:'Impossible de valider un compte dont l’adresse e-mail n’est pas vérifiée.'});
            db.run(`UPDATE users SET status='approved' WHERE id=?`, [userId], function(err) {
                if (err) return res.status(500).json({error:'Erreur lors de la validation.'});
                return res.json({message:'Le compte a été validé.'});
            });
        });
    }

    let newStatus = action === 'approve' ? 'approved' : action === 'disable' ? 'disabled' : 'rejected';

    db.run(`UPDATE users SET status = ? WHERE id = ?`, [newStatus, userId], function(err) {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la mise à jour du statut du compte.' });
        }

        if (this.changes === 0) {
            return res.status(404).json({ error: 'Utilisateur non trouvé.' });
        }

        res.json({ message: `Le compte a été ${newStatus === 'approved' ? 'validé' : 'refusé'} avec succès.` });
    });
});

app.post('/api/auth/login', (req, res) => {
    const { username, password } = req.body;
    const loginUsername = String(username || '').trim();

    if (!loginUsername || !password) {
        return res.status(400).json({ error: 'Veuillez entrer le nom d\'utilisateur et le mot de passe.' });
    }

    db.get(`SELECT * FROM users WHERE LOWER(TRIM(username)) = LOWER(TRIM(?))`, [loginUsername], async (err, user) => {
        if (err || !user) {
            return res.status(401).json({ error: 'Nom d\'utilisateur ou mot de passe incorrect.' });
        }

        const isPasswordValid = await bcrypt.compare(password, user.password);
        if (!isPasswordValid) return res.status(401).json({ error: 'Nom d\'utilisateur ou mot de passe incorrect.' });
        if (Number(user.email_verified) !== 1) return res.status(403).json({error:'Adresse e-mail non vérifiée.'});
        if (Number(user.must_change_password) === 1) return res.status(403).json({error:'Changez votre mot de passe temporaire avant de vous connecter.'});
        if (!sameText(user.status, 'approved')) return res.status(403).json({ error: `Votre compte n'est pas encore validé. Statut actuel : ${user.status}` });

        db.get(`SELECT id, is_validator, max_groupes_l3, max_groupes_4ing, max_groupes_m2_5ing, max_groupes_m2, max_groupes_5ing, max_groupes_m2_5ing_mixte
                FROM enseignants_import WHERE LOWER(TRIM(email)) = LOWER(TRIM(?))`, [user.email], (teacherErr, teacherRow) => {
        const sessionUser = {
            id: user.id, username: user.username, role: user.role,
            nom: user.nom, prenom: user.prenom, email: user.email,
            niveau: user.niveau, specialite: user.specialite, matricule: user.matricule,
            teacher_id: teacherRow ? teacherRow.id : null,
            is_validator: teacherRow ? teacherRow.is_validator : 0,
            max_groupes_l3: teacherRow ? teacherRow.max_groupes_l3 : 0,
            max_groupes_4ing: teacherRow ? teacherRow.max_groupes_4ing : 0,
            max_groupes_m2_5ing: teacherRow ? teacherRow.max_groupes_m2_5ing : 2,
            max_groupes_m2: teacherRow ? teacherRow.max_groupes_m2 : 0,
            max_groupes_5ing: teacherRow ? teacherRow.max_groupes_5ing : 0,
            max_groupes_m2_5ing_mixte: teacherRow ? teacherRow.max_groupes_m2_5ing_mixte : 0
        };
        const token = crypto.randomBytes(32).toString('hex');
        activeSessions.set(token, sessionUser);
        res.json({
            message: 'Connexion réussie.',
            token,
            user: {
                ...sessionUser
            }
        });
        });
    });
});

app.post('/api/auth/logout', requireAuth, (req, res) => {
    const authorization = req.headers.authorization || '';
    const token = authorization.startsWith('Bearer ') ? authorization.slice(7) : null;
    if (token) activeSessions.delete(token);
    res.json({ message: 'Déconnexion réussie.' });
});

// ==========================================
// ROUTE API - CRÉATION DES GROUPES D'ÉTUDIANTS
// ==========================================
app.post('/api/groupes/creer', requireAuth, (req, res, next) => {
    if (!sameText(req.user.role, 'student')) return res.status(403).json({ error: 'Seuls les étudiants peuvent créer un groupe.' });
    next();
}, (req, res, next) => {
    phaseState('groupes', (err, state) => err ? res.status(500).json({error:'Vérification de la période impossible.'}) : !state.open ? res.status(403).json({error:'La période de formation des groupes est fermée.'}) : next());
}, uploadSujet.single('file'), (req, res) => {
    const { annee_universitaire, niveau, specialite, etudiants_ids, sujet_titre } = req.body;
    const annee = String(annee_universitaire || '').trim();
    const niveauGroupe = canonicalLevel(niveau);
    const specialiteDemandee = String(specialite || '').trim();

    if (!annee || !niveauGroupe || !etudiants_ids) {
        return res.status(400).json({ error: 'Informations manquantes pour créer le groupe.' });
    }

    let identifiants;
    try {
        identifiants = typeof etudiants_ids === 'string' ? JSON.parse(etudiants_ids) : etudiants_ids;
    } catch {
        return res.status(400).json({ error: 'Format des identifiants des étudiants invalide.' });
    }
    if (!Array.isArray(identifiants)) {
        return res.status(400).json({ error: 'La liste des autres membres est invalide.' });
    }

    // Le créateur connecté est ajouté automatiquement comme premier membre.
    if (!req.user.matricule) return res.status(400).json({ error: 'Votre compte étudiant ne possède pas de matricule associé.' });
    const membresSaisis = [String(req.user.matricule).trim(), ...identifiants.map(v => String(v).trim()).filter(Boolean)];
    const uniques = [...new Set(membresSaisis.map(normalizeMatricule))];
    if (uniques.length !== membresSaisis.length) {
        // Le matricule du créateur peut être omis du champ ; s'il y est saisi, on évite le doublon.
        const sansCreateur = membresSaisis.slice(1);
        if (sansCreateur.some(v => normalizeMatricule(v) === normalizeMatricule(req.user.matricule))) {
            return res.status(400).json({ error: 'Votre propre matricule est déjà ajouté automatiquement : ne le ressaisissez pas.' });
        }
        return res.status(400).json({ error: 'Les matricules des autres membres doivent être distincts.' });
    }

    const taille = membresSaisis.length;
    let type_groupe;
    if (sameText(niveauGroupe, 'L3') && taille === 3) type_groupe = 'trinome';
    else if (sameText(niveauGroupe, '4ING') && taille >= 4 && taille <= 6) type_groupe = 'groupe_4_6';
    else if ((sameText(niveauGroupe, 'Master 2') || sameText(niveauGroupe, '5ING')) && taille === 2) type_groupe = 'binome';
    else if (sameText(niveauGroupe, 'M2 + 5ING') && taille === 2) type_groupe = 'binome_mixte_m2_5ing';
    else {
        const regle = sameText(niveauGroupe, 'L3') ? 'exactement 3 étudiants' :
            sameText(niveauGroupe, '4ING') ? 'entre 4 et 6 étudiants' :
            'exactement 2 étudiants';
        return res.status(400).json({ error: `Règle ${niveauGroupe} : le groupe doit contenir ${regle}.` });
    }

    const placeholders = membresSaisis.map(() => '?').join(',');
    const lookupQuery = `
        SELECT id, matricule, annee_universitaire, niveau, specialite
        FROM etudiants_import
        WHERE LOWER(TRIM(matricule)) IN (${placeholders}) OR CAST(id AS TEXT) IN (${placeholders})
    `;
    db.all(lookupQuery, [...membresSaisis.map(normalizeMatricule), ...membresSaisis], (lookupErr, students) => {
        if (lookupErr) return res.status(500).json({ error: 'Erreur lors de la vérification des étudiants.' });
        const byMatricule = new Map(students.map(st => [normalizeMatricule(st.matricule), st]));
        const byId = new Map(students.map(st => [String(st.id), st]));
        const resolved = membresSaisis.map(token => byMatricule.get(normalizeMatricule(token)) || byId.get(token));
        if (resolved.some(st => !st)) return res.status(400).json({ error: 'Un ou plusieurs matricules ne correspondent à aucun étudiant importé.' });
        if (new Set(resolved.map(st => st.id)).size !== resolved.length) return res.status(400).json({ error: 'Un même étudiant ne peut pas être ajouté plusieurs fois au groupe.' });

        const wrongYear = resolved.find(st => !sameText(st.annee_universitaire, annee));
        if (wrongYear) return res.status(400).json({ error: 'Tous les membres doivent appartenir à la même année universitaire.' });
        if (sameText(niveauGroupe, 'L3') && resolved.some(st => !sameText(canonicalLevel(st.niveau), 'L3'))) return res.status(400).json({ error: 'Un groupe L3 doit être composé exclusivement d’étudiants L3.' });
        if (sameText(niveauGroupe, '4ING') && resolved.some(st => !sameText(canonicalLevel(st.niveau), '4ING'))) return res.status(400).json({ error: 'Un groupe 4ING doit être composé exclusivement d’étudiants 4ING.' });
        if (sameText(niveauGroupe, 'Master 2') && resolved.some(st => !sameText(canonicalLevel(st.niveau), 'Master 2'))) return res.status(400).json({ error: 'Un binôme Master 2 doit contenir deux étudiants de Master 2.' });
        if (sameText(niveauGroupe, '5ING') && resolved.some(st => !sameText(canonicalLevel(st.niveau), '5ING'))) return res.status(400).json({ error: 'Un binôme 5ING doit contenir deux étudiants de 5ING.' });
        if (sameText(niveauGroupe, 'M2 + 5ING') && !(resolved.some(st => sameText(canonicalLevel(st.niveau), 'Master 2')) && resolved.some(st => sameText(canonicalLevel(st.niveau), '5ING')))) {
            return res.status(400).json({ error: 'Un binôme M2 + 5ING doit contenir exactement un étudiant de Master 2 et un étudiant de 5ING.' });
        }
        if (sameText(niveauGroupe, '4ING') && resolved.some(st => !sameText(st.specialite, specialiteDemandee))) {
            return res.status(400).json({ error: 'Les étudiants 4ING doivent appartenir à la spécialité sélectionnée.' });
        }
        // En L3, le regroupement inter-spécialités est autorisé.
        const specialitesMembres = [...new Map(resolved.map(st => [normalizeText(st.specialite), String(st.specialite || '').trim()]).filter(([key, value]) => key && value)).values()];
        const specialiteGroupe = sameText(niveauGroupe, 'L3') ? (specialitesMembres.length > 1 ? 'Mixte' : (specialitesMembres[0] || 'Toutes')) : sameText(niveauGroupe, 'M2 + 5ING') ? 'Mixte' : (specialiteDemandee || specialitesMembres[0] || 'Toutes');
        const ids = resolved.map(st => st.id);
        const idPlaceholders = ids.map(() => '?').join(',');
        db.get(
            `SELECT gm.etudiant_id FROM groupe_membres gm
             JOIN groupes g ON gm.groupe_id = g.id
             WHERE LOWER(TRIM(g.annee_universitaire)) = LOWER(TRIM(?)) AND gm.etudiant_id IN (${idPlaceholders})`,
            [annee, ...ids],
            (checkErr, existingMember) => {
                if (checkErr) return res.status(500).json({ error: 'Erreur lors de la vérification des membres.' });
                if (existingMember) return res.status(400).json({ error: 'L’un des étudiants sélectionnés appartient déjà à un autre groupe.' });

                const uploadedPath = req.file ? req.file.path : null;
                db.serialize(() => {
                    db.run('BEGIN IMMEDIATE TRANSACTION', (beginErr) => {
                        if (beginErr) return res.status(500).json({ error: 'Impossible de démarrer la transaction.' });
                        const statutInitial = uploadedPath ? 'en_attente' : 'valide';
                        db.run(
                            `INSERT INTO groupes
                             (annee_universitaire, niveau, specialite, type_groupe, sujet_titre, sujet_description_path, statut_sujet)
                             VALUES (?, ?, ?, ?, ?, ?, ?)`,
                            [annee, niveauGroupe, specialiteGroupe, type_groupe, sujet_titre || null, uploadedPath, statutInitial],
                            function(groupErr) {
                                if (groupErr) {
                                    return db.run('ROLLBACK', () => res.status(500).json({ error: 'Erreur lors de la création du groupe.' }));
                                }
                                const groupeId = this.lastID;
                                let remaining = ids.length;
                                let failed = false;
                                ids.forEach((etudiantId, index) => {
                                    db.run(
                                        `INSERT INTO groupe_membres (groupe_id, etudiant_id, role_dans_groupe) VALUES (?, ?, ?)`,
                                        [groupeId, etudiantId, index === 0 ? 'chef_groupe' : 'membre'],
                                        (memberErr) => {
                                            if (memberErr && !failed) {
                                                failed = true;
                                                return db.run('ROLLBACK', () => res.status(500).json({ error: 'Erreur lors de l’association des membres au groupe.' }));
                                            }
                                            remaining--;
                                            if (remaining === 0 && !failed) {
                                                db.run('COMMIT', (commitErr) => {
                                                    if (commitErr) return db.run('ROLLBACK', () => res.status(500).json({ error: 'Erreur lors de la validation de la transaction.' }));
                                                    res.json({ message: 'Groupe créé avec succès !', groupe_id: groupeId, type_groupe, nombre_membres: taille });
                                                });
                                            }
                                        }
                                    );
                                });
                            }
                        );
                    });
                });
            }
        );
    });
});

// ==========================================
// ROUTES API - GESTION DES QUOTAS ET SUJETS INTERNES (ENSEIGNANTS)
// ==========================================

// 1. Mise à jour des quotas d'encadrement (Minimum 2 binômes exigé pour M2/5ING)
app.post('/api/enseignants/:id/quotas', (req, res) => {
    const enseignantId = req.params.id;
    let { max_groupes_l3, max_groupes_4ing, max_groupes_m2_5ing } = req.body;

    max_groupes_l3 = parseInt(max_groupes_l3) || 0;
    max_groupes_4ing = parseInt(max_groupes_4ing) || 0;
    max_groupes_m2_5ing = parseInt(max_groupes_m2_5ing) || 2;
    const max_groupes_m2 = Math.max(0, parseInt(req.body.max_groupes_m2, 10) || 0);
    const max_groupes_5ing = Math.max(0, parseInt(req.body.max_groupes_5ing, 10) || 0);
    const max_groupes_m2_5ing_mixte = Math.max(0, parseInt(req.body.max_groupes_m2_5ing_mixte, 10) || 0);

    // Règle pédagogique : Minimum 2 binômes en M2/5ING
    if (max_groupes_m2_5ing < 2) {
        return res.status(400).json({ error: 'Règle pédagogique : Chaque enseignant doit encadrer au minimum 2 binômes en Master 2 / 5ING.' });
    }

    const query = `
        UPDATE enseignants_import 
        SET max_groupes_l3 = ?, max_groupes_4ing = ?, max_groupes_m2_5ing = ?, max_groupes_m2=?, max_groupes_5ing=?, max_groupes_m2_5ing_mixte=?
        WHERE id = ?
    `;

    db.run(query, [max_groupes_l3, max_groupes_4ing, max_groupes_m2_5ing, max_groupes_m2, max_groupes_5ing, max_groupes_m2_5ing_mixte, enseignantId], function(err) {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la mise à jour des quotas.' });
        }

        if (this.changes === 0) {
            return res.status(404).json({ error: 'Enseignant non trouvé.' });
        }

        res.json({ message: 'Quotas d\'encadrement mis à jour avec succès (Minimum de 2 binômes M2/5ING respecté).' });
    });
});

// 2. Dépôt d'un sujet interne par un enseignant
app.post('/api/enseignants/sujets/creer', (req, res) => {
    const { annee_universitaire, enseignant_id, niveau, specialite, titre, description, capacite_max } = req.body;
    const anneeSujet = String(annee_universitaire || '').trim();
    const niveauSujet = canonicalLevel(niveau);
    const specialiteSujet = String(specialite || 'Toutes').trim();
    const titreSujet = String(titre || '').trim();

    if (!anneeSujet || !enseignant_id || !niveauSujet || !titreSujet) {
        return res.status(400).json({ error: 'Veuillez remplir les champs obligatoires (année, enseignant, niveau, titre).' });
    }

    const capacity = Number.parseInt(capacite_max, 10);
    if (!Number.isInteger(capacity) || capacity < 1) return res.status(400).json({ error: 'La capacité maximale doit être un entier positif.' });
    db.get('SELECT id FROM enseignants_import WHERE id=?', [enseignant_id], (teacherErr, teacher) => {
      if (teacherErr || !teacher) return res.status(400).json({ error: 'Enseignant invalide.' });
      db.get(`SELECT COUNT(*) total FROM etudiants_import WHERE LOWER(TRIM(annee_universitaire))=LOWER(TRIM(?)) AND (LOWER(TRIM(niveau))=LOWER(TRIM(?)) OR (LOWER(TRIM(?))='m2 + 5ing' AND LOWER(TRIM(niveau)) IN ('master 2','5ing'))) AND (LOWER(TRIM(?))='toutes' OR LOWER(TRIM(specialite))=LOWER(TRIM(?)))`, [anneeSujet, niveauSujet, niveauSujet, specialiteSujet, specialiteSujet], (checkErr, eligible) => {
        if (checkErr) return res.status(500).json({ error: 'Vérification du niveau/spécialité impossible.' });
        if (!eligible.total) return res.status(400).json({ error: 'Aucun étudiant importé ne correspond à cette année, ce niveau et cette spécialité.' });
        const query = `INSERT INTO sujets_internes (annee_universitaire, enseignant_id, niveau, specialite, titre, description, capacite_max) VALUES (?, ?, ?, ?, ?, ?, ?)`;
        db.run(query, [anneeSujet, enseignant_id, niveauSujet, specialiteSujet, titreSujet, String(description || '').trim(), capacity], function(err) {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de l\'enregistrement du sujet.' });
        }

        res.json({
            message: 'Sujet interne créé avec succès.',
            sujet_id: this.lastID
        });
      });
    });
    });
});


app.get('/api/groupes/:id/fichier-sujet', requireAuth, (req, res) => {
    db.get('SELECT sujet_description_path FROM groupes WHERE id = ?', [req.params.id], (err, row) => {
        if (err) return res.status(500).json({ error: 'Erreur lors de la recherche du document.' });
        if (!row || !row.sujet_description_path) return res.status(404).json({ error: 'Document introuvable.' });
        const authorizeDownload = callback => {
            if (sameText(req.user.role, 'admin') || sameText(req.user.role, 'validator') || Number(req.user.is_validator) === 1) return callback(true);
            if (!sameText(req.user.role, 'student') || !req.user.matricule) return callback(false);
            db.get(`SELECT 1 ok FROM groupe_membres gm JOIN etudiants_import st ON st.id=gm.etudiant_id WHERE gm.groupe_id=? AND LOWER(TRIM(st.matricule))=LOWER(TRIM(?))`, [req.params.id, req.user.matricule], (memberErr, member) => callback(!memberErr && !!member));
        };
        const filePath = path.resolve(row.sujet_description_path);
        const allowedRoot = path.resolve(sujetsDir) + path.sep;
        if (!filePath.startsWith(allowedRoot) || !fs.existsSync(filePath)) return res.status(404).json({ error: 'Document introuvable.' });
        authorizeDownload(allowed => allowed ? res.download(filePath) : res.status(403).json({error:'Vous ne disposez pas des droits pour consulter ce document.'}));
    });
});

// ==========================================
// ROUTES API - VALIDATION DES SUJETS EXTERNES (COMITÉ / VALIDATEURS)
// ==========================================
// Les routes de validation à votes indépendants sont déclarées dans les extensions ci-dessous.

// ==========================================
// ROUTE API - ATTRIBUTION DES SUJETS ET ENCADRANTS AUX GROUPES
// ==========================================
app.post('/api/admin/groupes/:id/affecter', (req,res) => {
  const groupeId=req.params.id; const {sujet_id,enseignant_id}=req.body;
  if(!enseignant_id)return res.status(400).json({error:'Veuillez spécifier l’enseignant encadrant.'});
  db.get('SELECT * FROM groupes WHERE id=?',[groupeId],(e,g)=>{
    if(e||!g)return res.status(404).json({error:'Groupe non trouvé.'});
    db.get('SELECT * FROM enseignants_import WHERE id=?',[enseignant_id],(e2,t)=>{
      if(e2||!t)return res.status(404).json({error:'Enseignant non trouvé.'});
      const quotaCol=quotaField(g.niveau);
      let max=quotaCol?Number(t[quotaCol]||0):0;
      if(!max && ['Master 2','5ING','M2 + 5ING'].includes(g.niveau))max=Number(t.max_groupes_m2_5ing||0);
      db.get('SELECT COUNT(*) total FROM groupes WHERE enseignant_id=? AND LOWER(TRIM(niveau))=LOWER(TRIM(?)) AND id<>?',[enseignant_id,g.niveau,groupeId],(e3,c)=>{
        if(e3)return res.status(500).json({error:'Vérification de capacité impossible.'});
        if(c.total>=max)return res.status(400).json({error:`Capacité atteinte pour ${g.niveau} (${c.total}/${max}).`});
        const finish=(topic)=>{
          const sql=`UPDATE groupes SET enseignant_id=?,sujet_titre=COALESCE(?,sujet_titre),sujet_interne_id=?,statut_sujet=CASE WHEN ? IS NOT NULL THEN 'valide' ELSE statut_sujet END WHERE id=?`;
          db.run(sql,[enseignant_id,topic?topic.titre:null,topic?topic.id:null,topic?topic.id:null,groupeId],err=>err?res.status(500).json({error:'Erreur lors de l’affectation.'}):res.json({message:'Affectation enregistrée.',groupe_id:Number(groupeId),enseignant:`${t.prenom} ${t.nom}`,niveau:g.niveau,charge_actuelle:c.total+1,max_autorise:max}));
        };
        if(!sujet_id)return finish(null);
        db.get(`SELECT * FROM sujets_internes WHERE id=? AND LOWER(TRIM(annee_universitaire))=LOWER(TRIM(?)) AND (LOWER(TRIM(niveau))=LOWER(TRIM(?)) OR LOWER(TRIM(niveau))='toutes') AND (LOWER(TRIM(specialite))='toutes' OR LOWER(TRIM(specialite))=LOWER(TRIM(?)))`,[sujet_id,g.annee_universitaire,g.niveau,g.specialite],(e4,topic)=>{
          if(e4)return res.status(500).json({error:'Vérification du sujet impossible.'});
          if(!topic)return res.status(400).json({error:'Le sujet interne ne correspond pas à l’année, au niveau ou à la spécialité du groupe.'});
          db.get('SELECT COUNT(*) total FROM groupes WHERE sujet_interne_id=? AND id<>?',[sujet_id,groupeId],(e5,count)=>{
            if(e5)return res.status(500).json({error:'Vérification de la capacité du sujet impossible.'});
            if(count.total>=Number(topic.capacite_max))return res.status(400).json({error:`Capacité maximale du sujet atteinte (${topic.capacite_max} groupes).`});
            finish(topic);
          });
        });
      });
    });
  });
});

// ==========================================
// ROUTES API - CONSULTATION (DASHBOARDS & LISTES)
// ==========================================
app.get('/api/sujets-internes', (req, res) => {
    const { niveau, specialite } = req.query;
    let query = `
        SELECT s.*, 
               e.nom as enseignant_nom, e.prenom as enseignant_prenom
        FROM sujets_internes s
        JOIN enseignants_import e ON s.enseignant_id = e.id
        WHERE 1=1
    `;
    const params = [];

    if (niveau) {
        query += ` AND LOWER(TRIM(s.niveau)) = LOWER(TRIM(?))`;
        params.push(String(niveau).trim());
    }
    if (specialite) {
        query += ` AND (LOWER(TRIM(s.specialite)) = LOWER(TRIM(?)) OR LOWER(TRIM(s.specialite)) = 'toutes')`;
        params.push(String(specialite).trim());
    }

    db.all(query, params, (err, rows) => {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la récupération des sujets.' });
        }
        res.json({ sujets_internes: rows });
    });
});

app.get('/api/admin/users/pending', (req, res) => {
    db.all(`SELECT id, username, email, nom, prenom, role, matricule, niveau, specialite, email_verified, created_at FROM users WHERE status = 'pending'`, [], (err, rows) => {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la récupération des utilisateurs en attente.' });
        }
        res.json({ users_pending: rows });
    });
});



// ==========================================
// ROUTES API - GESTION ADMINISTRATIVE (AJOUTS MANUELS & VALIDATEURS)
// ==========================================

// 1. Ajout manuel d'un étudiant
app.post('/api/admin/etudiants/ajouter', (req, res) => {
    const { annee_universitaire, niveau, specialite, matricule, nom, prenom, section, groupe } = req.body;
    const anneeClean = String(annee_universitaire || '').trim();

    if (!annee_universitaire || !niveau || !specialite || !matricule || !nom || !prenom) {
        return res.status(400).json({ error: 'Veuillez remplir tous les champs obligatoires de l\'étudiant.' });
    }

    const query = `
        INSERT OR IGNORE INTO etudiants_import 
        (annee_universitaire, niveau, specialite, matricule, nom, prenom, section, groupe) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `;

    db.run(query, [anneeClean, canonicalLevel(niveau), String(specialite).trim(), String(matricule).trim(), String(nom).trim(), String(prenom).trim(), String(section || '').trim(), String(groupe || '').trim()], function(err) {
        if (err) return res.status(500).json({ error: 'Erreur lors de l’ajout de l’étudiant.' });
        if (this.changes === 0) return res.status(409).json({ error: 'Ce matricule existe déjà.' });
        res.json({ message: 'Étudiant ajouté manuellement avec succès !' });
    });
});

// 2. Ajout manuel d'un enseignant
app.post('/api/admin/enseignants/ajouter', (req, res) => {
    const { annee_universitaire, nom, prenom, email } = req.body;

    if (!annee_universitaire || !nom || !prenom || !email) {
        return res.status(400).json({ error: 'Veuillez remplir tous les champs obligatoires de l\'enseignant.' });
    }

    const query = `
        INSERT OR IGNORE INTO enseignants_import 
        (annee_universitaire, nom, prenom, email, max_groupes_m2_5ing) 
        VALUES (?, ?, ?, ?, 2)
    `;

    db.run(query, [annee_universitaire, nom, prenom, email], function(err) {
        if (err) return res.status(500).json({ error: 'Erreur lors de l’ajout de l’enseignant.' });
        if (this.changes === 0) return res.status(409).json({ error: 'Cet email existe déjà.' });
        res.json({ message: 'Enseignant ajouté manuellement avec succès !' });
    });
});

// 3. Basculer le rôle de validateur pour un enseignant
app.post('/api/admin/enseignants/:id/toggle-validator', (req, res) => {
    const enseignantId = req.params.id;
    const { is_validator } = req.body; // 1 ou 0

    const query = `UPDATE enseignants_import SET is_validator = ? WHERE id = ?`;

    db.run(query, [is_validator ? 1 : 0, enseignantId], function(err) {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la modification du statut de validateur.' });
        }
        if (this.changes === 0) {
            return res.status(404).json({ error: 'Enseignant non trouvé.' });
        }
        res.json({ message: `Statut de validateur mis à jour avec succès.` });
    });
});

// 4. Récupérer la liste complète des enseignants (avec leur statut de validateur)
app.get('/api/admin/enseignants', (req, res) => {
    db.all(`SELECT * FROM enseignants_import`, [], (err, rows) => {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la récupération des enseignants.' });
        }
        res.json({ enseignants: rows });
    });
});

// ==========================================
// ROUTES API - GESTION DES NIVEAUX ET SPÉCIALITÉS
// ==========================================

// 1. Ajouter un niveau et sa spécialité
app.post('/api/admin/specialites/ajouter', (req, res) => {
    const { niveau, nom_specialite } = req.body;
    if (!niveau || !nom_specialite) {
        return res.status(400).json({ error: 'Le niveau et le nom de la spécialité sont obligatoires.' });
    }

    const query = `INSERT INTO specialites (niveau, nom_specialite) VALUES (?, ?)`;
    db.run(query, [niveau, nom_specialite], function(err) {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de l\'ajout de la spécialité.' });
        }
        res.json({ message: 'Spécialité ajoutée avec succès !' });
    });
});

// 2. Obtenir tous les niveaux et spécialités
app.get('/api/admin/specialites', (req, res) => {
    db.all(`SELECT * FROM specialites ORDER BY niveau, nom_specialite`, [], (err, rows) => {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la récupération des spécialités.' });
        }
        res.json({ specialites: rows });
    });
});

// ==========================================
// ROUTES API - CONSULTATION ÉTUDIANTS (FILTRES + RECHERCHE)
// ==========================================
app.get('/api/admin/etudiants/recherche', (req, res) => {
    const { niveau, specialite, q } = req.query;

    let query = `SELECT * FROM etudiants_import WHERE 1=1`;
    const params = [];

    if (niveau) {
        query += ` AND LOWER(TRIM(niveau)) = LOWER(TRIM(?))`;
        params.push(String(niveau).trim());
    }
    if (specialite) {
        query += ` AND LOWER(TRIM(specialite)) = LOWER(TRIM(?))`;
        params.push(String(specialite).trim());
    }
    if (q) {
        query += ` AND (nom LIKE ? OR prenom LIKE ? OR matricule LIKE ?)`;
        const searchTerm = `%${q}%`;
        params.push(searchTerm, searchTerm, searchTerm);
    }

    query += ` ORDER BY nom, prenom`;

    db.all(query, params, (err, rows) => {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la recherche des étudiants.' });
        }
        res.json({ etudiants: rows });
    });
});

// ==========================================
// ROUTES API - ENSEIGNANTS ET LEURS GROUPES ENCADRÉS
// ==========================================
app.get('/api/admin/enseignants/:id/groupes', (req, res) => {
    const enseignantId = req.params.id;

    const query = `
        SELECT g.*, 
               GROUP_CONCAT(e.nom || ' ' || e.prenom, ', ') as membres
        FROM groupes g
        JOIN groupe_membres gm ON g.id = gm.groupe_id
        JOIN etudiants_import e ON gm.etudiant_id = e.id
        WHERE g.enseignant_id = ?
        GROUP BY g.id
    `;

    db.all(query, [enseignantId], (err, rows) => {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la récupération des groupes de l\'enseignant.' });
        }
        res.json({ groupes: rows });
    });
});

// ==========================================
// ROUTES API - MODIFICATION ET SUPPRESSION DES SPÉCIALITÉS
// ==========================================

// 1. Modifier une spécialité existante
app.put('/api/admin/specialites/:id', (req, res) => {
    const { id } = req.params;
    const { niveau, nom_specialite } = req.body;

    if (!niveau || !nom_specialite) {
        return res.status(400).json({ error: 'Le niveau et le nom de la spécialité sont obligatoires.' });
    }

    const query = `UPDATE specialites SET niveau = ?, nom_specialite = ? WHERE id = ?`;
    db.run(query, [niveau, nom_specialite, id], function(err) {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la modification de la spécialité.' });
        }
        if (this.changes === 0) {
            return res.status(404).json({ error: 'Spécialité non trouvée.' });
        }
        res.json({ message: 'Spécialité mise à jour avec succès !' });
    });
});

// 2. Supprimer une spécialité
app.delete('/api/admin/specialites/:id', (req, res) => {
    const { id } = req.params;

    const query = `DELETE FROM specialites WHERE id = ?`;
    db.run(query, [id], function(err) {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la suppression de la spécialité.' });
        }
        if (this.changes === 0) {
            return res.status(404).json({ error: 'Spécialité non trouvée.' });
        }
        res.json({ message: 'Spécialité supprimée avec succès !' });
    });
});

// ==========================================
// MODIFICATION & SUPPRESSION : ÉTUDIANTS & ENSEIGNANTS
// ==========================================

// --- ÉTUDIANTS ---

// Modifier un étudiant
app.put('/api/admin/etudiants/:id', (req, res) => {
    const { id } = req.params;
    const { niveau, specialite, matricule, nom, prenom, section, groupe } = req.body;

    const query = `UPDATE etudiants_import SET niveau = ?, specialite = ?, matricule = ?, nom = ?, prenom = ?, section = ?, groupe = ? WHERE id = ?`;
    db.run(query, [niveau, specialite, matricule, nom, prenom, section, groupe, id], function(err) {
        if (err) return res.status(500).json({ error: 'Erreur lors de la modification de l\'étudiant.' });
        if (this.changes === 0) return res.status(404).json({ error: 'Étudiant non trouvé.' });
        res.json({ message: 'Étudiant mis à jour avec succès !' });
    });
});

// Supprimer un étudiant
app.delete('/api/admin/etudiants/:id', (req, res) => {
    const { id } = req.params;
    db.run(`DELETE FROM etudiants_import WHERE id = ?`, [id], function(err) {
        if (err) return res.status(500).json({ error: 'Erreur lors de la suppression de l\'étudiant.' });
        if (this.changes === 0) return res.status(404).json({ error: 'Étudiant non trouvé.' });
        res.json({ message: 'Étudiant supprimé avec succès !' });
    });
});


// --- ENSEIGNANTS ---

// Modifier un enseignant
app.put('/api/admin/enseignants/:id', (req, res) => {
    const { id } = req.params;
    const { nom, prenom, email } = req.body;

    const query = `UPDATE enseignants_import SET nom = ?, prenom = ?, email = ? WHERE id = ?`;
    db.run(query, [nom, prenom, email, id], function(err) {
        if (err) return res.status(500).json({ error: 'Erreur lors de la modification de l\'enseignant.' });
        if (this.changes === 0) return res.status(404).json({ error: 'Enseignant non trouvé.' });
        res.json({ message: 'Enseignant mis à jour avec succès !' });
    });
});

// Supprimer un enseignant
app.delete('/api/admin/enseignants/:id', (req, res) => {
    const { id } = req.params;
    db.run(`DELETE FROM enseignants_import WHERE id = ?`, [id], function(err) {
        if (err) return res.status(500).json({ error: 'Erreur lors de la suppression de l\'enseignant.' });
        if (this.changes === 0) return res.status(404).json({ error: 'Enseignant non trouvé.' });
        res.json({ message: 'Enseignant supprimé avec succès !' });
    });
});

// ==========================================
// ROUTES ADMIN : GESTION DES GROUPES PFE
// ==========================================

// 1. Récupérer tous les groupes (avec filtre optionnel par niveau) et leurs membres
app.get('/api/admin/groupes', (req, res) => {
    const { niveau } = req.query;
    
    let query = `
        SELECT g.*, 
               t.nom as enseignant_nom, t.prenom as enseignant_prenom,
               GROUP_CONCAT(e.nom || ' ' || e.prenom || ' (' || e.matricule || ')', ', ') as membres
        FROM groupes g
        LEFT JOIN enseignants_import t ON g.enseignant_id = t.id
        JOIN groupe_membres gm ON g.id = gm.groupe_id
        JOIN etudiants_import e ON gm.etudiant_id = e.id
    `;
    
    const params = [];
    if (niveau) {
        query += ` WHERE LOWER(TRIM(g.niveau)) = LOWER(TRIM(?))`;
        params.push(String(niveau).trim());
    }
    
    query += ` GROUP BY g.id ORDER BY g.niveau, g.id DESC`;

    db.all(query, params, (err, rows) => {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la récupération des groupes.' });
        }
        res.json({ groupes: rows });
    });
});

// 2. Supprimer un groupe (et libérer ses membres)
app.delete('/api/admin/groupes/:id', (req, res) => {
    const { id } = req.params;

    // Supprimer d'abord les associations de membres dans groupe_membres
    db.run(`DELETE FROM groupe_membres WHERE groupe_id = ?`, [id], (err) => {
        if (err) {
            return res.status(500).json({ error: 'Erreur lors de la suppression des membres du groupe.' });
        }

        // Ensuite supprimer le groupe lui-même
        db.run(`DELETE FROM groupes WHERE id = ?`, [id], function(err) {
            if (err) {
                return res.status(500).json({ error: 'Erreur lors de la suppression du groupe.' });
            }
            if (this.changes === 0) {
                return res.status(404).json({ error: 'Groupe non trouvé.' });
            }
            res.json({ message: 'Groupe supprimé avec succès !' });
        });
    });
});


// ============================================================================
// EXTENSIONS CAHIER DES CHARGES : PHASES, VALIDATIONS, PREFERENCES, AFFECTATIONS
// ============================================================================
const schemaMigrations = [
  ['users','email_verified','INTEGER NOT NULL DEFAULT 0'],
  ['users','email_verification_code_hash','TEXT'],
  ['users','email_verification_expires','DATETIME'],
  ['users','temporary_password_expires','DATETIME'],
  ['users','must_change_password','INTEGER NOT NULL DEFAULT 1'],
  ['enseignants_import','max_groupes_m2','INTEGER NOT NULL DEFAULT 2'],
  ['enseignants_import','max_groupes_5ing','INTEGER NOT NULL DEFAULT 2'],
  ['enseignants_import','max_groupes_m2_5ing_mixte','INTEGER NOT NULL DEFAULT 2'],
  ['groupes','sujet_interne_id','INTEGER']
];
function runSchemaMigrations(index=0) {
  if (index >= schemaMigrations.length) {
    // Les comptes déjà présents avant cette version restent utilisables.
    return db.run(`UPDATE users SET email_verified=1,must_change_password=0
                   WHERE temporary_password_expires IS NULL AND email_verification_code_hash IS NULL`, [], err => {
      if (err) console.error('Migration des comptes existants :', err.message);
      createFeatureTables();
    });
  }
  const [table,column,definition] = schemaMigrations[index];
  db.all(`PRAGMA table_info(${table})`, [], (err, cols) => {
    if (err) return console.error(`Migration ${table}.${column}:`,err.message);
    if (cols.some(c=>c.name===column)) return runSchemaMigrations(index+1);
    db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`, e => {
      if (e) console.error(`Migration ${table}.${column}:`,e.message);
      runSchemaMigrations(index+1);
    });
  });
}
function createFeatureTables() {
  db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS phases_projet (
      id INTEGER PRIMARY KEY AUTOINCREMENT, phase_key TEXT NOT NULL UNIQUE,
      libelle TEXT NOT NULL, date_debut TEXT, date_fin TEXT, updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS validations_sujets (
      id INTEGER PRIMARY KEY AUTOINCREMENT, groupe_id INTEGER NOT NULL, validateur_id INTEGER NOT NULL,
      decision TEXT NOT NULL CHECK(decision IN ('valide','invalide')), commentaire TEXT,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP, UNIQUE(groupe_id, validateur_id),
      FOREIGN KEY(groupe_id) REFERENCES groupes(id) ON DELETE CASCADE,
      FOREIGN KEY(validateur_id) REFERENCES enseignants_import(id) ON DELETE CASCADE
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS preferences_enseignants (
      id INTEGER PRIMARY KEY AUTOINCREMENT, enseignant_id INTEGER NOT NULL, groupe_id INTEGER NOT NULL,
      rang INTEGER DEFAULT 1, created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(enseignant_id,groupe_id), FOREIGN KEY(enseignant_id) REFERENCES enseignants_import(id) ON DELETE CASCADE,
      FOREIGN KEY(groupe_id) REFERENCES groupes(id) ON DELETE CASCADE
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS campagnes_affectation (
      id INTEGER PRIMARY KEY AUTOINCREMENT, annee_universitaire TEXT NOT NULL, categorie TEXT NOT NULL,
      statut TEXT NOT NULL DEFAULT 'brouillon' CHECK(statut IN ('brouillon','validee')),
      resultat_json TEXT NOT NULL, created_at DATETIME DEFAULT CURRENT_TIMESTAMP, validated_at DATETIME
    )`);
    db.run(`INSERT OR IGNORE INTO phases_projet(phase_key,libelle) VALUES
      ('groupes','Formation des groupes et dépôt des sujets externes'),
      ('validation','Validation des sujets externes'),
      ('preferences','Préférences des enseignants et affectations')`);
    db.run('SELECT 1', [], err => {
      if (err) return console.error('Initialisation des nouvelles tables :',err.message);
      schemaReady = true;
      console.log('Migrations et tables complémentaires prêtes.');
    });
  });
}
runSchemaMigrations();

function phaseState(key, cb) {
    db.get('SELECT * FROM phases_projet WHERE phase_key = ?', [key], (err, row) => {
        if (err) return cb(err);
        if (!row || (!row.date_debut && !row.date_fin)) return cb(null, { open: false, configured: false, phase: row || null });
        if (!row.date_debut || !row.date_fin) return cb(null, { open: false, configured: false, phase: row });
        const now = new Date();
        const start = new Date(row.date_debut);
        const end = new Date(row.date_fin);
        cb(null, { open: now >= start && now <= end, configured: true, phase: row });
    });
}
function csvEscape(value) { return '"' + String(value == null ? '' : value).replace(/"/g, '""') + '"'; }
function sendCsv(res, filename, headers, rows) {
    const body = '\uFEFF' + [headers, ...rows].map(r => r.map(csvEscape).join(';')).join('\r\n');
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(body);
}

app.get('/api/specialites', (req,res) => db.all(`SELECT niveau,nom_specialite FROM specialites UNION SELECT DISTINCT niveau,specialite AS nom_specialite FROM etudiants_import ORDER BY niveau,nom_specialite`, [], (e,rows) => e ? res.status(500).json({error:'Lecture des spécialités impossible.'}) : res.json({specialites:rows})));

// Consultation et configuration des trois phases par l'administrateur.
app.get('/api/admin/phases', (req,res) => db.all('SELECT * FROM phases_projet ORDER BY id', [], (e, rows) => e ? res.status(500).json({error:'Lecture des phases impossible.'}) : res.json({phases:rows})));
app.put('/api/admin/phases/:key', (req,res) => {
    const allowed = ['groupes','validation','preferences'];
    const key = req.params.key;
    const {date_debut,date_fin} = req.body;
    if (!allowed.includes(key)) return res.status(404).json({error:'Phase inconnue.'});
    if ((date_debut && isNaN(Date.parse(date_debut))) || (date_fin && isNaN(Date.parse(date_fin)))) return res.status(400).json({error:'Dates invalides.'});
    if (date_debut && date_fin && new Date(date_debut) > new Date(date_fin)) return res.status(400).json({error:'La date de début doit précéder la date de fin.'});
    db.run(`UPDATE phases_projet SET date_debut=?,date_fin=?,updated_at=CURRENT_TIMESTAMP WHERE phase_key=?`,[date_debut||null,date_fin||null,key],function(e){
        if(e) return res.status(500).json({error:'Mise à jour impossible.'});
        res.json({message:'Période mise à jour.',changes:this.changes});
    });
});
app.get('/api/phases', (req,res) => db.all('SELECT phase_key,libelle,date_debut,date_fin FROM phases_projet ORDER BY id',[],(e,rows)=>e?res.status(500).json({error:'Erreur.'}):res.json({phases:rows})));

app.get('/api/groupes/mes-groupes', requireAuth, (req,res)=>{
  if(!sameText(req.user.role,'student')) return res.status(403).json({error:'Accès réservé aux étudiants.'});
  db.all(`SELECT g.id,g.annee_universitaire,g.niveau,g.specialite,g.type_groupe,g.sujet_titre,g.sujet_description_path,g.statut_sujet,
    GROUP_CONCAT(st.prenom||' '||st.nom,' / ') membres
    FROM groupes g JOIN groupe_membres gm ON gm.groupe_id=g.id JOIN etudiants_import st ON st.id=gm.etudiant_id
    WHERE g.id IN (SELECT gm2.groupe_id FROM groupe_membres gm2 JOIN etudiants_import s2 ON s2.id=gm2.etudiant_id WHERE LOWER(TRIM(s2.matricule))=LOWER(TRIM(?)))
    GROUP BY g.id ORDER BY g.id DESC`,[req.user.matricule],(e,rows)=>e?res.status(500).json({error:'Lecture des groupes impossible.'}):res.json({groupes:rows}));
});

// Edition/suppression du sujet externe : autorisées uniquement pendant la phase 1.
app.put('/api/groupes/:id/sujet-externe', requireAuth, uploadSujet.single('file'), (req,res) => {
    if (!sameText(req.user.role, 'student')) return res.status(403).json({error:'Action réservée aux étudiants.'});
    phaseState('groupes',(e,state)=>{
        if(e) return res.status(500).json({error:'Vérification de la période impossible.'});
        if(!state.open) return res.status(403).json({error:'La période de dépôt/modification des sujets est fermée.'});
        db.get(`SELECT g.id,g.sujet_description_path FROM groupes g JOIN groupe_membres gm ON gm.groupe_id=g.id JOIN etudiants_import st ON st.id=gm.etudiant_id WHERE g.id=? AND LOWER(TRIM(st.matricule))=LOWER(TRIM(?))`,[req.params.id,req.user.matricule],(err,g)=>{
            if(err) return res.status(500).json({error:'Erreur base de données.'});
            if(!g) return res.status(404).json({error:'Groupe introuvable ou vous n’en êtes pas membre.'});
            const title = (req.body.sujet_titre || '').trim();
            const filePath = req.file ? req.file.path : g.sujet_description_path;
            if(req.file && g.sujet_description_path && fs.existsSync(g.sujet_description_path)) fs.unlink(g.sujet_description_path,()=>{});
            db.run('DELETE FROM validations_sujets WHERE groupe_id=?',[g.id],eVote=>{if(eVote)return res.status(500).json({error:'Réinitialisation des avis impossible.'});
              db.run(`UPDATE groupes SET sujet_titre=?,sujet_description_path=?,statut_sujet=? WHERE id=?`,[title||null,filePath,filePath?'en_attente':'valide',g.id],e2=>e2?res.status(500).json({error:'Modification impossible.'}):res.json({message:'Sujet externe mis à jour. Les avis précédents ont été réinitialisés.'}));
            });
        });
    });
});
app.delete('/api/groupes/:id/sujet-externe', requireAuth, (req,res) => {
    if(!sameText(req.user.role,'student')) return res.status(403).json({error:'Action réservée aux étudiants.'});
    phaseState('groupes',(e,state)=>{
        if(e) return res.status(500).json({error:'Erreur.'});
        if(!state.open) return res.status(403).json({error:'La période de modification des sujets est fermée.'});
        db.get(`SELECT g.id,g.sujet_description_path FROM groupes g JOIN groupe_membres gm ON gm.groupe_id=g.id JOIN etudiants_import st ON st.id=gm.etudiant_id WHERE g.id=? AND LOWER(TRIM(st.matricule))=LOWER(TRIM(?))`,[req.params.id,req.user.matricule],(err,g)=>{
            if(err) return res.status(500).json({error:'Erreur.'}); if(!g) return res.status(404).json({error:'Groupe introuvable.'});
            if(g.sujet_description_path && fs.existsSync(g.sujet_description_path)) fs.unlink(g.sujet_description_path,()=>{});
            db.run('DELETE FROM validations_sujets WHERE groupe_id=?',[g.id],eVote=>{if(eVote)return res.status(500).json({error:'Réinitialisation des avis impossible.'});
              db.run(`UPDATE groupes SET sujet_titre=NULL,sujet_description_path=NULL,statut_sujet='valide' WHERE id=?`,[g.id],e2=>e2?res.status(500).json({error:'Suppression impossible.'}):res.json({message:'Sujet externe supprimé.'}));
            });
        });
    });
});

// Votes individuels des validateurs : deux avis, puis troisième validateur en cas de désaccord.
app.get('/api/validateur/sujets-externes/attente', (req,res) => {
    const sql=`SELECT g.id groupe_id,g.annee_universitaire,g.niveau,g.specialite,g.type_groupe,g.sujet_titre,g.sujet_description_path,g.statut_sujet,
      (SELECT COUNT(*) FROM validations_sujets v WHERE v.groupe_id=g.id) nb_avis,
      (SELECT GROUP_CONCAT(v.decision, ', ') FROM validations_sujets v WHERE v.groupe_id=g.id) avis,
      e.nom chef_nom,e.prenom chef_prenom,e.matricule chef_matricule
      FROM groupes g JOIN groupe_membres gm ON gm.groupe_id=g.id AND gm.role_dans_groupe='chef_groupe'
      JOIN etudiants_import e ON e.id=gm.etudiant_id WHERE g.sujet_description_path IS NOT NULL AND g.statut_sujet='en_attente' ORDER BY g.id`;
    db.all(sql,[],(e,rows)=>e?res.status(500).json({error:'Erreur de lecture.'}):res.json({sujets_externes:rows}));
});
app.post('/api/validateur/sujets-externes/:id/arbitrage', (req,res,next) => phaseState('validation',(e,state) => e ? res.status(500).json({error:'Vérification de la période impossible.'}) : !state.open ? res.status(403).json({error:'La période de validation est fermée.'}) : next()), (req,res) => {
    const {action,commentaire} = req.body;
    if(!['valider','rejeter'].includes(action)) return res.status(400).json({error:'Décision invalide.'});
    const vote = action==='valider'?'valide':'invalide';
    db.get('SELECT id,statut_sujet FROM groupes WHERE id=? AND sujet_description_path IS NOT NULL',[req.params.id],(e,g)=>{
        if(e) return res.status(500).json({error:'Erreur.'}); if(!g) return res.status(404).json({error:'Sujet externe introuvable.'});
        db.all('SELECT validateur_id,decision FROM validations_sujets WHERE groupe_id=?',[g.id],(e2,votes)=>{
            if(e2) return res.status(500).json({error:'Erreur.'});
            const isThird = votes.length>=2 && votes[0].decision!==votes[1].decision;
            if(votes.length>=2 && !isThird) return res.status(409).json({error:'Les deux premiers avis concordent : le sujet est déjà décidé.'});
            if(isThird && votes.some(v=>Number(v.validateur_id)===Number(req.user.teacher_id))) return res.status(403).json({error:'Le troisième avis doit provenir d’un troisième validateur distinct.'});
            if(!isThird && votes.some(v=>Number(v.validateur_id)===Number(req.user.teacher_id))) return res.status(409).json({error:'Vous avez déjà donné votre avis sur ce sujet.'});
            if(votes.length>=3) return res.status(409).json({error:'Le nombre maximal de trois avis est atteint.'});
            db.run('INSERT INTO validations_sujets(groupe_id,validateur_id,decision,commentaire) VALUES(?,?,?,?)',[g.id,req.user.teacher_id,vote,commentaire||null],function(insErr){
                if(insErr) return res.status(409).json({error:'Votre avis n’a pas pu être enregistré.'});
                const all=[...votes,{validateur_id:req.user.teacher_id,decision:vote}];
                let final=null;
                if(all.length===2 && all[0].decision===all[1].decision) final=all[0].decision;
                if(all.length===3) final=vote; // avis du troisième arbitre en cas de désaccord
                const status=final || 'en_attente';
                db.run('UPDATE groupes SET statut_sujet=? WHERE id=?',[status,g.id],e3=>e3?res.status(500).json({error:'Avis enregistré, mais statut non actualisé.'}):res.json({message:final?`Décision finale : ${final==='valide'?'sujet validé':'sujet rejeté'}.`:'Avis enregistré. Un second avis indépendant est attendu.',nb_avis:all.length,statut:status}));
            });
        });
    });
});

// Préférences des enseignants : consultation filtrée et sélection des groupes.
app.get('/api/enseignants/groupes-disponibles', (req,res)=>{
    const {niveau}=req.query; let sql=`SELECT g.id,g.annee_universitaire,g.niveau,g.specialite,g.type_groupe,CASE WHEN g.statut_sujet='valide' THEN g.sujet_titre ELSE NULL END sujet_titre,g.statut_sujet,
      GROUP_CONCAT(st.prenom||' '||st.nom,' / ') membres FROM groupes g JOIN groupe_membres gm ON gm.groupe_id=g.id
      JOIN etudiants_import st ON st.id=gm.etudiant_id WHERE g.enseignant_id IS NULL AND (g.sujet_description_path IS NULL OR g.statut_sujet IN ('valide','invalide'))`;
    const params=[]; if(niveau){sql+=' AND LOWER(TRIM(g.niveau))=LOWER(TRIM(?))';params.push(String(niveau).trim());} sql+=' GROUP BY g.id ORDER BY g.id';
    db.all(sql,params,(e,rows)=>e?res.status(500).json({error:'Erreur.'}):res.json({groupes:rows}));
});
app.post('/api/enseignants/preferences', (req,res)=>{
    phaseState('preferences',(e,state)=>{
        if(e) return res.status(500).json({error:'Erreur.'}); if(!state.open) return res.status(403).json({error:'La période de préférences est fermée.'});
        const ids=Array.isArray(req.body.groupe_ids)?[...new Set(req.body.groupe_ids.map(Number).filter(Number.isInteger))]:[];
        db.serialize(()=>{
            db.run('BEGIN IMMEDIATE');
            db.run('DELETE FROM preferences_enseignants WHERE enseignant_id=?',[req.user.teacher_id],err=>{
                if(err) return db.run('ROLLBACK',()=>res.status(500).json({error:'Erreur.'}));
                let i=0; const next=()=>{if(i>=ids.length)return db.run('COMMIT',()=>res.json({message:'Préférences enregistrées.'})); const id=ids[i++]; db.run('INSERT INTO preferences_enseignants(enseignant_id,groupe_id,rang) SELECT ?,id,? FROM groupes WHERE id=? AND enseignant_id IS NULL',[req.user.teacher_id,i,id],e2=>e2?db.run('ROLLBACK',()=>res.status(500).json({error:'Préférence invalide.'})):next());}; next();
            });
        });
    });
});

// Capacités explicites par catégorie (les anciennes colonnes sont conservées pour compatibilité).
app.put('/api/admin/enseignants/:id/capacites', (req,res)=>{
    const fields=['max_groupes_l3','max_groupes_4ing','max_groupes_m2','max_groupes_5ing','max_groupes_m2_5ing_mixte'];
    const vals=fields.map(k=>Math.max(0,parseInt(req.body[k],10)||0));
    db.run(`UPDATE enseignants_import SET ${fields.map(k=>k+'=?').join(',')} WHERE id=?`,[...vals,req.params.id],function(e){e?res.status(500).json({error:'Mise à jour impossible.'}):this.changes?res.json({message:'Capacités mises à jour.'}):res.status(404).json({error:'Enseignant introuvable.'});});
});

function quotaField(level){const key=canonicalLevel(level);return ({'L3':'max_groupes_l3','4ING':'max_groupes_4ing','Master 2':'max_groupes_m2','5ING':'max_groupes_5ing','M2 + 5ING':'max_groupes_m2_5ing_mixte'})[key]||null;}
app.post('/api/admin/affectations/lancer', (req,res)=>{
    const {annee_universitaire,niveau}=req.body; const anneeCampagne=String(annee_universitaire || '').trim(); const niveauCampagne=canonicalLevel(niveau); const field=quotaField(niveauCampagne);
    if(!anneeCampagne||!field) return res.status(400).json({error:'Année ou catégorie invalide.'});
    db.all(`SELECT g.id,g.niveau,g.specialite,g.statut_sujet,g.sujet_titre FROM groupes g WHERE LOWER(TRIM(g.annee_universitaire))=LOWER(TRIM(?)) AND LOWER(TRIM(g.niveau))=LOWER(TRIM(?)) AND g.enseignant_id IS NULL AND (g.sujet_description_path IS NULL OR g.statut_sujet IN ('valide','invalide')) ORDER BY g.id`,[anneeCampagne,niveauCampagne],(e,groups)=>{
        if(e) return res.status(500).json({error:'Lecture des groupes impossible.'});
        db.all(`SELECT t.id,t.nom,t.prenom,t.${field} capacite,(SELECT COUNT(*) FROM groupes gx WHERE gx.enseignant_id=t.id AND LOWER(TRIM(gx.niveau))=LOWER(TRIM(?))) deja_affectes FROM enseignants_import t ORDER BY t.nom,t.prenom`,[niveauCampagne],(e2,teachers)=>{
            if(e2) return res.status(500).json({error:'Lecture des enseignants impossible.'});
            db.all(`SELECT p.enseignant_id,p.groupe_id,p.rang FROM preferences_enseignants p JOIN groupes g ON g.id=p.groupe_id WHERE LOWER(TRIM(g.annee_universitaire))=LOWER(TRIM(?)) AND LOWER(TRIM(g.niveau))=LOWER(TRIM(?)) ORDER BY p.rang`,[anneeCampagne,niveauCampagne],(e3,prefs)=>{
                if(e3) return res.status(500).json({error:'Lecture des préférences impossible.'});
                const remaining=Object.fromEntries(teachers.map(t=>[t.id,Math.max(0,(t.capacite||0)-(t.deja_affectes||0))]));
                const loadCounts=Object.fromEntries(teachers.map(t=>[t.id,Number(t.deja_affectes||0)]));
                const assignments=[]; const unassigned=[];
                const sorted=[...groups].sort((a,b)=>{const pa=prefs.filter(p=>p.groupe_id===a.id).length,pb=prefs.filter(p=>p.groupe_id===b.id).length;return pb-pa;});
                sorted.forEach(g=>{
                  const choices=prefs.filter(p=>p.groupe_id===g.id).sort((a,b)=>a.rang-b.rang).map(p=>p.enseignant_id);
                  const candidates=teachers.filter(t=>remaining[t.id]>0).sort((a,b)=>{
                    const rankA=choices.indexOf(a.id),rankB=choices.indexOf(b.id);
                    const scoreA=loadCounts[a.id]/Math.max(1,a.capacite||1)+(rankA>=0?rankA*0.25:0.75);
                    const scoreB=loadCounts[b.id]/Math.max(1,b.capacite||1)+(rankB>=0?rankB*0.25:0.75);
                    return scoreA-scoreB || a.nom.localeCompare(b.nom);
                  });
                  const t=candidates[0];
                  if(t){remaining[t.id]--;loadCounts[t.id]++;assignments.push({groupe_id:g.id,enseignant_id:t.id,enseignant:t.prenom+' '+t.nom});}
                  else unassigned.push({groupe_id:g.id,raison:'Capacité disponible insuffisante'});
                });
                const result={annee_universitaire:anneeCampagne,niveau:niveauCampagne,affectations:assignments,non_affectes:unassigned,capacites_restantes:remaining,enseignants:teachers.map(t=>({id:t.id,nom:t.nom,prenom:t.prenom,capacite:t.capacite||0}))};
                db.run(`INSERT INTO campagnes_affectation(annee_universitaire,categorie,statut,resultat_json) VALUES(?,?,'brouillon',?)`,[anneeCampagne,niveauCampagne,JSON.stringify(result)],function(e4){if(e4)return res.status(500).json({error:'Enregistrement du brouillon impossible.'});res.json({campagne_id:this.lastID,...result});});
            });
        });
    });
});
app.post('/api/admin/affectations/:id/valider', (req,res)=>{
    db.get('SELECT * FROM campagnes_affectation WHERE id=? AND statut=\'brouillon\'',[req.params.id],(e,c)=>{
        if(e)return res.status(500).json({error:'Erreur.'});if(!c)return res.status(404).json({error:'Brouillon introuvable.'});
        const r=JSON.parse(c.resultat_json); db.serialize(()=>{db.run('BEGIN IMMEDIATE');let i=0;const next=()=>{if(i>=r.affectations.length)return db.run(`UPDATE campagnes_affectation SET statut='validee',validated_at=CURRENT_TIMESTAMP WHERE id=?`,[c.id],e2=>e2?db.run('ROLLBACK',()=>res.status(500).json({error:'Validation impossible.'})):db.run('COMMIT',()=>res.json({message:'Affectations validées.',affectes:r.affectations.length,non_affectes:r.non_affectes.length})));const a=r.affectations[i++];db.run('UPDATE groupes SET enseignant_id=? WHERE id=? AND enseignant_id IS NULL',[a.enseignant_id,a.groupe_id],function(e3){if(e3||this.changes!==1)return db.run('ROLLBACK',()=>res.status(409).json({error:'Un groupe a changé depuis la génération du brouillon. La validation a été annulée.'}));next();});};next();});
    });
});
app.put('/api/admin/affectations/:id/ajuster', (req,res)=>{
  const groupeId=Number(req.body.groupe_id), enseignantId=Number(req.body.enseignant_id);
  if(!Number.isInteger(groupeId)||!Number.isInteger(enseignantId))return res.status(400).json({error:'Groupe ou enseignant invalide.'});
  db.get("SELECT * FROM campagnes_affectation WHERE id=? AND statut='brouillon'",[req.params.id],(e,c)=>{
    if(e)return res.status(500).json({error:'Erreur.'});if(!c)return res.status(404).json({error:'Brouillon introuvable ou déjà validé.'});
    const result=JSON.parse(c.resultat_json);const groupExists=result.affectations.some(x=>x.groupe_id===groupeId)||result.non_affectes.some(x=>x.groupe_id===groupeId);
    if(!groupExists)return res.status(400).json({error:'Ce groupe ne figure pas dans ce brouillon.'});
    const field=quotaField(c.categorie);
    db.get(`SELECT t.id,t.nom,t.prenom,t.${field||'max_groupes_l3'} capacite,(SELECT COUNT(*) FROM groupes gx WHERE gx.enseignant_id=t.id AND LOWER(TRIM(gx.niveau))=LOWER(TRIM(?))) deja_affectes FROM enseignants_import t WHERE t.id=?`,[c.categorie,enseignantId],(e2,t)=>{
      if(e2)return res.status(500).json({error:'Erreur.'});if(!t)return res.status(404).json({error:'Enseignant introuvable.'});
      const used=result.affectations.filter(x=>x.enseignant_id===enseignantId&&x.groupe_id!==groupeId).length + Number(t.deja_affectes||0);
      if(used>=Number(t.capacite||0))return res.status(400).json({error:`Capacité atteinte pour cet enseignant (${used}/${t.capacite||0}). Ajustez d’abord sa capacité.`});
      result.affectations=result.affectations.filter(x=>x.groupe_id!==groupeId);
      result.non_affectes=result.non_affectes.filter(x=>x.groupe_id!==groupeId);
      result.affectations.push({groupe_id:groupeId,enseignant_id:enseignantId,enseignant:`${t.prenom} ${t.nom}`,ajustement_manuel:true});
      result.affectations.sort((a,b)=>a.groupe_id-b.groupe_id);
      db.run('UPDATE campagnes_affectation SET resultat_json=? WHERE id=?',[JSON.stringify(result),c.id],err=>err?res.status(500).json({error:'Enregistrement de l’ajustement impossible.'}):res.json({message:'Affectation du brouillon ajustée.',resultat:result}));
    });
  });
});
app.get('/api/admin/affectations/:id', (req,res)=>db.get('SELECT * FROM campagnes_affectation WHERE id=?',[req.params.id],(e,r)=>e?res.status(500).json({error:'Erreur.'}):r?res.json({...r,resultat:JSON.parse(r.resultat_json)}):res.status(404).json({error:'Campagne introuvable.'})));

app.get('/api/admin/statistiques/groupes',(req,res)=>{
  db.all(`SELECT g.niveau categorie,COUNT(DISTINCT g.id) nb_groupes,COUNT(gm.id) nb_etudiants,
    COUNT(DISTINCT CASE WHEN g.enseignant_id IS NOT NULL THEN g.id END) nb_groupes_affectes,
    COUNT(DISTINCT CASE WHEN g.sujet_description_path IS NOT NULL AND g.statut_sujet='valide' THEN g.id END) nb_sujets_externes_valides
    FROM groupes g LEFT JOIN groupe_membres gm ON gm.groupe_id=g.id GROUP BY g.niveau ORDER BY g.niveau`,[],(e,rows)=>e?res.status(500).json({error:'Calcul des statistiques impossible.'}):res.json({statistiques:rows}));
});

// Exports CSV par niveau/catégorie.
app.get('/api/admin/export/groupes', (req,res)=>{
    const {niveau}=req.query; let sql=`SELECT g.annee_universitaire,g.niveau,g.specialite,g.type_groupe,g.sujet_titre,g.statut_sujet,st.matricule,st.nom,st.prenom,st.specialite specialite_etudiant FROM groupes g JOIN groupe_membres gm ON gm.groupe_id=g.id JOIN etudiants_import st ON st.id=gm.etudiant_id`;
    const params=[]; if(niveau){sql+=' WHERE LOWER(TRIM(g.niveau))=LOWER(TRIM(?))';params.push(String(niveau).trim());} sql+=' ORDER BY g.niveau,g.id,gm.id';
    db.all(sql,params,(e,rows)=>e?res.status(500).json({error:'Export impossible.'}):sendCsv(res,'groupes.csv',['Année','Catégorie','Spécialité groupe','Type','Sujet','Statut sujet','Matricule','Nom','Prénom','Spécialité étudiant'],rows.map(r=>[r.annee_universitaire,r.niveau,r.specialite,r.type_groupe,r.sujet_titre,r.statut_sujet,r.matricule,r.nom,r.prenom,r.specialite_etudiant])));
});
app.get('/api/admin/export/enseignants', (req,res) => {
  db.all('SELECT nom,prenom,email,is_validator,max_groupes_l3,max_groupes_4ing,max_groupes_m2,max_groupes_5ing,max_groupes_m2_5ing_mixte FROM enseignants_import ORDER BY nom,prenom', [], (e, rows) => {
    if (e) return res.status(500).json({error:'Export impossible.'});
    sendCsv(res,'enseignants.csv',['Nom','Prénom','Email','Validateur','Capacité L3','Capacité 4ING','Capacité M2','Capacité 5ING','Capacité M2+5ING'],rows.map(r => [r.nom,r.prenom,r.email,r.is_validator,r.max_groupes_l3,r.max_groupes_4ing,r.max_groupes_m2,r.max_groupes_5ing,r.max_groupes_m2_5ing_mixte]));
  });
});
app.get('/api/admin/export/affectations', (req,res) => {
  db.all(`SELECT g.annee_universitaire,g.niveau,g.specialite,g.id groupe_id,e.nom,e.prenom,e.email FROM groupes g LEFT JOIN enseignants_import e ON e.id=g.enseignant_id ORDER BY g.niveau,g.id`, [], (e, rows) => {
    if (e) return res.status(500).json({error:'Export impossible.'});
    sendCsv(res,'affectations.csv',['Année','Catégorie','Spécialité','Groupe','Enseignant','Prénom enseignant','Email'],rows.map(r => [r.annee_universitaire,r.niveau,r.specialite,r.groupe_id,r.nom,r.prenom,r.email]));
  });
});

// Vidage sécurisé : liste blanche stricte et confirmation textuelle explicite.
app.post('/api/admin/maintenance/vider-tables',(req,res)=>{
    const allowed=['validations_sujets','preferences_enseignants','groupe_membres','groupes','sujets_internes','etudiants_import','enseignants_import','users','specialites','campagnes_affectation'];
    const tables=Array.isArray(req.body.tables)?[...new Set(req.body.tables)]:[];
    if(req.body.confirmation!=='VIDER LES DONNEES' || !tables.length || tables.some(t=>!allowed.includes(t))) return res.status(400).json({error:'Confirmation explicite ou liste de tables invalide.'});
    db.serialize(()=>{db.run('BEGIN IMMEDIATE');let i=0;const next=()=>{if(i>=tables.length)return db.run('COMMIT',e=>e?db.run('ROLLBACK',()=>res.status(500).json({error:'Transaction annulée.'})):res.json({message:'Tables vidées : '+tables.join(', ')}));db.run(`DELETE FROM ${tables[i++]}`,e=>e?db.run('ROLLBACK',()=>res.status(500).json({error:'Échec : transaction annulée.'})):next());};next();});
});


// Lancement du serveur
app.listen(PORT, () => {
    console.log(`Le serveur tourne sur http://localhost:${PORT}`);
});