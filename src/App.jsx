import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Users, GraduationCap, ClipboardList, Trophy, FileText, Building2,
  Settings, Download, Printer, Plus, Trash2, Pencil, X, Search,
  ChevronRight, School, Save, Loader2, AlertCircle, MessageSquare,
  Upload, Star, BarChart3, Send, UserCheck, CalendarDays, Database,
  WifiOff, LogOut, Lock, KeyRound, Eye, EyeOff
} from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";

/* ------------------------------------------------------------------ */
/*  Données de référence                                              */
/* ------------------------------------------------------------------ */

const CLASSES = [
  { id: "ps", name: "Petite Section", cycle: "Maternelle" },
  { id: "ms", name: "Moyenne Section", cycle: "Maternelle" },
  { id: "gs", name: "Grande Section", cycle: "Maternelle" },
  { id: "a1", name: "1ère année", cycle: "Primaire" },
  { id: "a2", name: "2ème année", cycle: "Primaire" },
  { id: "a3", name: "3ème année", cycle: "Primaire" },
  { id: "a4", name: "4ème année", cycle: "Primaire" },
  { id: "a5", name: "5ème année", cycle: "Primaire" },
  { id: "a6", name: "6ème année", cycle: "Secondaire" },
  { id: "a7", name: "7ème année", cycle: "Secondaire" },
  { id: "a8", name: "8ème année", cycle: "Secondaire" },
  { id: "a9", name: "9ème année", cycle: "Secondaire" },
  { id: "a10", name: "10ème année", cycle: "Secondaire" },
];

const DEFAULT_SUBJECTS = [
  { id: "fr", name: "Français", coef: 4 },
  { id: "math", name: "Mathématiques", coef: 4 },
  { id: "sci", name: "Sciences", coef: 2 },
  { id: "ang", name: "Anglais", coef: 2 },
  { id: "hg", name: "Histoire-Géographie", coef: 2 },
  { id: "eps", name: "EPS", coef: 1 },
  { id: "arts", name: "Arts / Activités", coef: 1 },
];

// Compte administrateur créé au tout premier démarrage de l'application.
// La direction peut ensuite ajouter d'autres comptes administrateurs
// depuis l'onglet Sécurité, et changer ce mot de passe à tout moment.
const SEED_ADMIN = {
  id: "admin-seed",
  name: "Bilguissou Diallo",
  email: "bilguissediallo22@gmail.com",
  password: "bilguissou26",
};

const STORAGE_KEY = "ecole_donnees_v1";
const MESSAGES_KEY = "ecole_messages_v1";
const NAME_KEY = "ecole_nom_utilisateur";

const DAYS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];

const TIME_SLOTS = [
  { id: "s1", label: "08h00 - 09h00" },
  { id: "s2", label: "09h00 - 10h00" },
  { id: "r1", label: "10h00 - 10h15", pause: "Récréation" },
  { id: "s3", label: "10h15 - 11h15" },
  { id: "s4", label: "11h15 - 12h15" },
  { id: "r2", label: "12h15 - 14h00", pause: "Pause déjeuner" },
  { id: "s5", label: "14h00 - 15h00" },
  { id: "s6", label: "15h00 - 16h00" },
];

const uid = () => Math.random().toString(36).slice(2, 10);

function mention(moy) {
  if (moy === null) return "—";
  if (moy >= 16) return "Excellent";
  if (moy >= 14) return "Très bien";
  if (moy >= 12) return "Bien";
  if (moy >= 10) return "Passable";
  return "Insuffisant";
}

/* ------------------------------------------------------------------ */
/*  Calculs                                                            */
/* ------------------------------------------------------------------ */

function computeAverage(studentId, grades, subjects) {
  const g = grades[studentId] || {};
  let totalPts = 0;
  let totalCoef = 0;
  subjects.forEach((s) => {
    const note = g[s.id];
    if (note !== undefined && note !== null && note !== "") {
      totalPts += parseFloat(note) * s.coef;
      totalCoef += s.coef;
    }
  });
  if (totalCoef === 0) return null;
  return Math.round((totalPts / totalCoef) * 100) / 100;
}

function rankClass(classId, students, grades, subjects) {
  const list = students
    .filter((st) => st.classId === classId)
    .map((st) => ({ ...st, moyenne: computeAverage(st.id, grades, subjects) }));

  const withNote = list.filter((s) => s.moyenne !== null).sort((a, b) => b.moyenne - a.moyenne);
  const withoutNote = list.filter((s) => s.moyenne === null);

  let rank = 0;
  let prevMoy = null;
  const ranked = withNote.map((s, i) => {
    if (s.moyenne !== prevMoy) rank = i + 1;
    prevMoy = s.moyenne;
    return { ...s, rang: rank };
  });

  return [...ranked, ...withoutNote.map((s) => ({ ...s, rang: null }))];
}

function studentsWithAverage(students, grades, subjects) {
  return students.map((s) => ({
    ...s,
    moyenne: computeAverage(s.id, grades, subjects),
    cls: CLASSES.find((c) => c.id === s.classId),
  }));
}

function normalizeText(str) {
  return (str || "")
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

/* ------------------------------------------------------------------ */
/*  App                                                                 */
/* ------------------------------------------------------------------ */

export default function EcoleApp() {
  const [loading, setLoading] = useState(true);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved | error
  const [schoolName, setSchoolName] = useState("Mon École");
  const [students, setStudents] = useState([]);
  const [subjects, setSubjects] = useState(DEFAULT_SUBJECTS);
  const [grades, setGrades] = useState({});
  const [teachers, setTeachers] = useState([]);
  const [schedules, setSchedules] = useState({});

  const [page, setPage] = useState("accueil");
  const [role, setRole] = useState(null); // null = pas encore connecté
  const [admins, setAdmins] = useState(null); // null = pas encore chargé
  const [currentUser, setCurrentUser] = useState(null);
  const [loginError, setLoginError] = useState("");
  const [activeClassId, setActiveClassId] = useState(CLASSES[3].id);
  const [activeSubjectId, setActiveSubjectId] = useState(DEFAULT_SUBJECTS[0].id);
  const [activeStudentId, setActiveStudentId] = useState(null);

  const [messages, setMessages] = useState([]);
  const [chatName, setChatNameState] = useState("");
  const [isOnline, setIsOnline] = useState(typeof navigator !== "undefined" ? navigator.onLine : true);

  useEffect(() => {
    const goOnline = () => setIsOnline(true);
    const goOffline = () => setIsOnline(false);
    window.addEventListener("online", goOnline);
    window.addEventListener("offline", goOffline);
    return () => {
      window.removeEventListener("online", goOnline);
      window.removeEventListener("offline", goOffline);
    };
  }, []);

  const firstLoad = useRef(true);
  const saveTimer = useRef(null);

  /* ---------- chargement initial ---------- */
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, true);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          setSchoolName(parsed.schoolName || "Mon École");
          setStudents(parsed.students || []);
          setSubjects(parsed.subjects && parsed.subjects.length ? parsed.subjects : DEFAULT_SUBJECTS);
          setGrades(parsed.grades || {});
          setTeachers(parsed.teachers || []);
          setSchedules(parsed.schedules || {});
          setAdmins(parsed.admins && parsed.admins.length ? parsed.admins : [SEED_ADMIN]);
        } else {
          setAdmins([SEED_ADMIN]);
        }
      } catch (e) {
        // pas encore de données -> on démarre avec la base vide
        setAdmins([SEED_ADMIN]);
      } finally {
        setLoading(false);
      }
      try {
        const nameRes = await window.storage.get(NAME_KEY, false);
        if (nameRes && nameRes.value) setChatNameState(nameRes.value);
      } catch (e) {
        // pas de nom enregistré
      }
    })();
  }, []);

  /* ---------- messagerie (partagée, avec actualisation périodique) ---------- */
  useEffect(() => {
    const loadMessages = async () => {
      try {
        const res = await window.storage.get(MESSAGES_KEY, true);
        if (res && res.value) setMessages(JSON.parse(res.value));
      } catch (e) {
        // pas encore de messages
      }
    };
    loadMessages();
    let interval;
    if (page === "chat") interval = setInterval(loadMessages, 4000);
    return () => interval && clearInterval(interval);
  }, [page]);

  const setChatName = async (name) => {
    setChatNameState(name);
    try {
      await window.storage.set(NAME_KEY, name, false);
    } catch (e) {
      // silencieux
    }
  };

  const sendMessage = async (text) => {
    const newMsg = { id: uid(), author: chatName, role, text, at: Date.now() };
    const updated = [...messages, newMsg];
    setMessages(updated);
    try {
      await window.storage.set(MESSAGES_KEY, JSON.stringify(updated), true);
    } catch (e) {
      // silencieux
    }
  };

  /* ---------- sauvegarde (debounce) ---------- */
  useEffect(() => {
    if (loading) return;
    if (firstLoad.current) { firstLoad.current = false; return; }
    setSaveState("saving");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        const payload = JSON.stringify({ schoolName, students, subjects, grades, teachers, schedules, admins });
        const res = await window.storage.set(STORAGE_KEY, payload, true);
        setSaveState(res ? "saved" : "error");
      } catch (e) {
        setSaveState("error");
      }
    }, 500);
    return () => clearTimeout(saveTimer.current);
  }, [schoolName, students, subjects, grades, teachers, schedules, admins, loading]);

  const studentsInClass = (classId) => students.filter((s) => s.classId === classId);

  /* ---------- actions élèves ---------- */
  const addStudent = (student) => setStudents((prev) => [...prev, { id: uid(), ...student }]);
  const updateStudent = (id, patch) =>
    setStudents((prev) => prev.map((s) => (s.id === id ? { ...s, ...patch } : s)));
  const deleteStudent = (id) => {
    setStudents((prev) => prev.filter((s) => s.id !== id));
    setGrades((prev) => {
      const copy = { ...prev };
      delete copy[id];
      return copy;
    });
  };
  const addManyStudents = (list) =>
    setStudents((prev) => [...prev, ...list.map((s) => ({ id: uid(), ...s }))]);

  /* ---------- actions enseignants ---------- */
  const addTeacher = (teacher) => setTeachers((prev) => [...prev, { id: uid(), ...teacher }]);
  const updateTeacher = (id, patch) =>
    setTeachers((prev) => prev.map((t) => (t.id === id ? { ...t, ...patch } : t)));
  const deleteTeacher = (id) => setTeachers((prev) => prev.filter((t) => t.id !== id));

  /* ---------- actions emploi du temps ---------- */
  const setScheduleCell = (cycle, key, value) => {
    setSchedules((prev) => ({
      ...prev,
      [cycle]: { ...(prev[cycle] || {}), [key]: value },
    }));
  };

  /* ---------- actions comptes administrateurs ---------- */
  const addAdmin = (admin) => setAdmins((prev) => [...prev, { id: uid(), ...admin }]);
  const removeAdmin = (id) => setAdmins((prev) => (prev.length > 1 ? prev.filter((a) => a.id !== id) : prev));

  /* ---------- actions notes ---------- */
  const setGrade = (studentId, subjectId, value) => {
    setGrades((prev) => ({
      ...prev,
      [studentId]: { ...(prev[studentId] || {}), [subjectId]: value },
    }));
  };

  /* ---------- actions matières ---------- */
  const addSubject = (subject) => setSubjects((prev) => [...prev, { id: uid(), ...subject }]);
  const updateSubject = (id, patch) =>
    setSubjects((prev) => prev.map((s) => (s.id === id ? { ...s, ...patch } : s)));
  const deleteSubject = (id) => setSubjects((prev) => prev.filter((s) => s.id !== id));

  /* ---------- export excel ---------- */
  const exportExcel = () => {
    const wb = XLSX.utils.book_new();

    const eleveRows = students.map((s) => {
      const cls = CLASSES.find((c) => c.id === s.classId);
      return {
        Nom: s.lastName,
        Prénom: s.firstName,
        Classe: cls ? cls.name : "",
        "Date de naissance": s.birthDate || "",
        Sexe: s.sexe || "",
        "Nom du parent/tuteur": s.parentName || "",
        Contact: s.parentContact || "",
      };
    });
    const wsEleves = XLSX.utils.json_to_sheet(eleveRows);
    XLSX.utils.book_append_sheet(wb, wsEleves, "Élèves");

    const noteRows = [];
    CLASSES.forEach((cls) => {
      const ranked = rankClass(cls.id, students, grades, subjects);
      ranked.forEach((s) => {
        const row = {
          Classe: cls.name,
          Nom: s.lastName,
          Prénom: s.firstName,
          Rang: s.rang ?? "",
        };
        subjects.forEach((subj) => {
          row[subj.name] = (grades[s.id] && grades[s.id][subj.id]) ?? "";
        });
        row["Moyenne générale"] = s.moyenne ?? "";
        row["Mention"] = mention(s.moyenne);
        noteRows.push(row);
      });
    });
    const wsNotes = XLSX.utils.json_to_sheet(noteRows);
    XLSX.utils.book_append_sheet(wb, wsNotes, "Notes et classement");

    const statRows = CLASSES.map((cls) => ({
      Classe: cls.name,
      Cycle: cls.cycle,
      "Effectif": studentsInClass(cls.id).length,
    }));
    const wsStats = XLSX.utils.json_to_sheet(statRows);
    XLSX.utils.book_append_sheet(wb, wsStats, "Effectifs");

    const teacherRows = teachers.map((t) => ({
      Nom: t.lastName,
      Prénom: t.firstName,
      "Matière(s)": t.subjectsTaught || "",
      "Classe(s) en charge": t.classesInCharge || "",
      Téléphone: t.phone || "",
      Email: t.email || "",
    }));
    const wsTeachers = XLSX.utils.json_to_sheet(teacherRows);
    XLSX.utils.book_append_sheet(wb, wsTeachers, "Enseignants");

    XLSX.writeFile(wb, `${schoolName.replace(/\s+/g, "_")}_donnees.xlsx`);
  };

  /* ---------- sauvegarde complète (.json) — pour transférer les données entre appareils ---------- */
  const exportBackup = () => {
    const payload = {
      version: 1,
      exportedAt: new Date().toISOString(),
      schoolName, students, subjects, grades, teachers, schedules, messages,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${schoolName.replace(/\s+/g, "_")}_sauvegarde_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const importBackup = (file) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        setSchoolName(data.schoolName || "Mon École");
        setStudents(data.students || []);
        setSubjects(data.subjects && data.subjects.length ? data.subjects : DEFAULT_SUBJECTS);
        setGrades(data.grades || {});
        setTeachers(data.teachers || []);
        setSchedules(data.schedules || {});
        setMessages(data.messages || []);
        window.storage.set(MESSAGES_KEY, JSON.stringify(data.messages || []), true).catch(() => {});
      } catch (e) {
        alert("Ce fichier de sauvegarde est illisible ou corrompu.");
      }
    };
    reader.readAsText(file);
  };

  const totalStudents = students.length;
  const activeClass = CLASSES.find((c) => c.id === activeClassId);
  const activeStudent = students.find((s) => s.id === activeStudentId);

  if (loading || admins === null) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", background: "#EDE4D3" }}>
        <Loader2 className="spin" size={28} color="#1B2A4A" />
      </div>
    );
  }

  if (!role) {
    return (
      <>
        <GlobalStyle />
        <LoginGate
          error={loginError}
          onLogin={(chosenRole, email, pwd) => {
            const cleanEmail = email.trim().toLowerCase();
            if (chosenRole === "direction") {
              const match = admins.find((a) => a.email.toLowerCase() === cleanEmail && a.password === pwd);
              if (match) {
                setLoginError("");
                setCurrentUser(match);
                setRole("direction");
                if (!chatName) setChatName(match.name);
              } else {
                setLoginError("Email ou mot de passe administrateur incorrect.");
              }
            } else {
              const match = teachers.find((t) => (t.email || "").toLowerCase() === cleanEmail && t.password && t.password === pwd);
              if (match) {
                setLoginError("");
                setCurrentUser(match);
                setRole("prof");
                if (!chatName) setChatName(`${match.firstName} ${match.lastName}`.trim());
              } else {
                setLoginError("Email ou mot de passe incorrect, ou aucun accès n'a encore été assigné par la direction.");
              }
            }
          }}
        />
      </>
    );
  }

  return (
    <div className="app">
      <GlobalStyle />
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-badge"><School size={18} /></div>
          <input
            className="brand-name"
            value={schoolName}
            onChange={(e) => setSchoolName(e.target.value)}
            aria-label="Nom de l'établissement"
          />
        </div>

        <div className="session-badge">
          <span>{currentUser?.name || `${currentUser?.firstName || ""} ${currentUser?.lastName || ""}`.trim() || (role === "direction" ? "Direction" : "Professeur")}</span>
          <button onClick={() => { setRole(null); setCurrentUser(null); setLoginError(""); }} title="Se déconnecter"><LogOut size={13} /></button>
        </div>

        <nav className="nav">
          <NavItem icon={<Building2 size={17} />} label="Accueil" active={page === "accueil"} onClick={() => setPage("accueil")} />
          <NavItem icon={<Users size={17} />} label="Élèves" active={page === "eleves"} onClick={() => setPage("eleves")} />
          <NavItem icon={<ClipboardList size={17} />} label="Notes" active={page === "notes"} onClick={() => setPage("notes")} />
          <NavItem icon={<Trophy size={17} />} label="Classement" active={page === "classement"} onClick={() => setPage("classement")} />
          <NavItem icon={<FileText size={17} />} label="Fiche élève" active={page === "fiche"} onClick={() => setPage("fiche")} />
          <NavItem icon={<CalendarDays size={17} />} label="Emploi du temps" active={page === "emploi"} onClick={() => setPage("emploi")} />
          <NavItem icon={<MessageSquare size={17} />} label="Discussion" active={page === "chat"} onClick={() => setPage("chat")} />
          {role === "direction" && (
            <>
              <NavItem icon={<UserCheck size={17} />} label="Enseignants" active={page === "enseignants"} onClick={() => setPage("enseignants")} />
              <NavItem icon={<GraduationCap size={17} />} label="Vue d'ensemble" active={page === "direction"} onClick={() => setPage("direction")} />
              <NavItem icon={<Star size={17} />} label="Palmarès" active={page === "palmares"} onClick={() => setPage("palmares")} />
              <NavItem icon={<BarChart3 size={17} />} label="Statistiques" active={page === "stats"} onClick={() => setPage("stats")} />
              <NavItem icon={<Upload size={17} />} label="Importer" active={page === "import"} onClick={() => setPage("import")} />
              <NavItem icon={<KeyRound size={17} />} label="Sécurité" active={page === "securite"} onClick={() => setPage("securite")} />
            </>
          )}
          <NavItem icon={<Settings size={17} />} label="Matières" active={page === "matieres"} onClick={() => setPage("matieres")} />
          <NavItem icon={<Database size={17} />} label="Sauvegarde" active={page === "sauvegarde"} onClick={() => setPage("sauvegarde")} />
        </nav>

        <div className="sidebar-foot">
          <div className={`conn-badge ${isOnline ? "is-online" : "is-offline"}`}>
            {isOnline ? "En ligne" : <><WifiOff size={11} /> Hors-ligne — les données restent sur cet appareil</>}
          </div>
          <button className="export-btn" onClick={exportExcel}>
            <Download size={15} /> Exporter en Excel
          </button>
          <div className="save-state">
            {saveState === "saving" && <>Enregistrement…</>}
            {saveState === "saved" && <>Enregistré</>}
            {saveState === "error" && <><AlertCircle size={12} /> Erreur d'enregistrement</>}
            {saveState === "idle" && <>&nbsp;</>}
          </div>
        </div>
      </aside>

      <main className="main">
        {page === "accueil" && (
          <AccueilPage
            schoolName={schoolName}
            totalStudents={totalStudents}
            studentsInClass={studentsInClass}
            setPage={setPage}
            setActiveClassId={setActiveClassId}
          />
        )}

        {page === "eleves" && (
          <ElevesPage
            activeClassId={activeClassId}
            setActiveClassId={setActiveClassId}
            students={studentsInClass(activeClassId)}
            addStudent={addStudent}
            updateStudent={updateStudent}
            deleteStudent={deleteStudent}
          />
        )}

        {page === "enseignants" && role === "direction" && (
          <EnseignantsPage
            teachers={teachers}
            addTeacher={addTeacher}
            updateTeacher={updateTeacher}
            deleteTeacher={deleteTeacher}
          />
        )}

        {page === "emploi" && (
          <EmploiDuTempsPage schedules={schedules} setScheduleCell={setScheduleCell} />
        )}

        {page === "notes" && (
          <NotesPage
            activeClassId={activeClassId}
            setActiveClassId={setActiveClassId}
            activeSubjectId={activeSubjectId}
            setActiveSubjectId={setActiveSubjectId}
            subjects={subjects}
            students={studentsInClass(activeClassId)}
            grades={grades}
            setGrade={setGrade}
          />
        )}

        {page === "classement" && (
          <ClassementPage
            activeClassId={activeClassId}
            setActiveClassId={setActiveClassId}
            students={students}
            grades={grades}
            subjects={subjects}
          />
        )}

        {page === "fiche" && (
          <FichePage
            students={students}
            activeStudentId={activeStudentId}
            setActiveStudentId={setActiveStudentId}
            grades={grades}
            subjects={subjects}
            schoolName={schoolName}
          />
        )}

        {page === "direction" && role === "direction" && (
          <DirectionPage students={students} studentsInClass={studentsInClass} grades={grades} subjects={subjects} />
        )}

        {page === "chat" && (
          <ChatPage messages={messages} sendMessage={sendMessage} chatName={chatName} setChatName={setChatName} role={role} />
        )}

        {page === "palmares" && role === "direction" && (
          <PalmaresPage students={students} grades={grades} subjects={subjects} />
        )}

        {page === "stats" && role === "direction" && (
          <StatistiquesPage students={students} grades={grades} subjects={subjects} studentsInClass={studentsInClass} />
        )}

        {page === "import" && role === "direction" && (
          <ImportPage addManyStudents={addManyStudents} />
        )}

        {page === "securite" && role === "direction" && (
          <SecuritePage admins={admins} addAdmin={addAdmin} removeAdmin={removeAdmin} currentUser={currentUser} />
        )}

        {page === "matieres" && (
          <MatieresPage subjects={subjects} addSubject={addSubject} updateSubject={updateSubject} deleteSubject={deleteSubject} />
        )}

        {page === "sauvegarde" && (
          <SauvegardePage exportBackup={exportBackup} importBackup={importBackup} exportExcel={exportExcel} isOnline={isOnline} />
        )}
      </main>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Nav item                                                           */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/*  Authentification : connexion                                       */
/* ------------------------------------------------------------------ */

function LoginGate({ error, onLogin }) {
  const [tab, setTab] = useState("prof");
  const [email, setEmail] = useState("");
  const [pwd, setPwd] = useState("");
  const [showPwd, setShowPwd] = useState(false);

  const submit = () => {
    if (!email.trim() || !pwd) return;
    onLogin(tab, email, pwd);
  };

  return (
    <div className="auth-screen">
      <div className="auth-card">
        <div className="brand-badge auth-badge"><School size={20} /></div>
        <h1>Connexion</h1>
        <p className="lede">Choisis ton espace, puis entre ton email et ton mot de passe.</p>

        <div className="mode-switch auth-tabs">
          <button className={tab === "prof" ? "active" : ""} onClick={() => setTab("prof")}>Professeur</button>
          <button className={tab === "direction" ? "active" : ""} onClick={() => setTab("direction")}>Direction</button>
        </div>

        <input
          className="auth-email"
          type="email"
          placeholder="Adresse email"
          value={email}
          autoFocus
          onChange={(e) => setEmail(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && submit()}
        />

        <div className="auth-pwd-row">
          <input
            type={showPwd ? "text" : "password"}
            placeholder="Mot de passe"
            value={pwd}
            onChange={(e) => setPwd(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && submit()}
          />
          <button type="button" className="auth-eye" onClick={() => setShowPwd((v) => !v)}>
            {showPwd ? <EyeOff size={16} /> : <Eye size={16} />}
          </button>
        </div>

        {error && <p className="auth-error"><AlertCircle size={13} /> {error}</p>}

        <button className="primary-btn auth-submit" type="button" onClick={submit}><Lock size={14} /> Se connecter</button>

        {tab === "prof" && (
          <p className="auth-note">Un professeur n'a accès que si la direction lui a créé un compte et un mot de passe.</p>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Sécurité (comptes administrateurs)                            */
/* ------------------------------------------------------------------ */

function SecuritePage({ admins, addAdmin, removeAdmin, currentUser }) {
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [pwd, setPwd] = useState("");
  const [error, setError] = useState("");
  const [saved, setSaved] = useState(false);

  const submit = () => {
    setSaved(false);
    if (!name.trim() || !email.trim()) {
      setError("Le nom et l'email sont obligatoires.");
      return;
    }
    if (pwd.length < 4) {
      setError("Le mot de passe doit contenir au moins 4 caractères.");
      return;
    }
    const cleanEmail = email.trim().toLowerCase();
    if (admins.some((a) => a.email.toLowerCase() === cleanEmail)) {
      setError("Un compte administrateur existe déjà avec cet email.");
      return;
    }
    addAdmin({ name: name.trim(), email: cleanEmail, password: pwd });
    setName(""); setEmail(""); setPwd(""); setError(""); setSaved(true);
  };

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Réservé à la direction</p>
        <h1>Sécurité — comptes administrateurs</h1>
        <p className="lede">
          Seuls les comptes de cette liste peuvent se connecter en tant que Direction : ajouter des enseignants et
          leur assigner un mot de passe, importer des élèves, consulter les statistiques, etc.
        </p>
      </header>

      <div className="ledger" style={{ marginBottom: 22 }}>
        <table>
          <thead><tr><th>Nom</th><th>Email</th><th></th></tr></thead>
          <tbody>
            {admins.map((a) => (
              <tr key={a.id}>
                <td>{a.name} {currentUser?.id === a.id && <span className="you-tag">(vous)</span>}</td>
                <td>{a.email}</td>
                <td className="row-actions">
                  <button
                    onClick={() => removeAdmin(a.id)}
                    className="danger"
                    title={admins.length === 1 ? "Impossible de supprimer le dernier compte admin" : "Supprimer"}
                    disabled={admins.length === 1}
                  >
                    <Trash2 size={14} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <h3 className="cycle-title">Ajouter un compte administrateur</h3>
      <div className="backup-card" style={{ maxWidth: 420 }}>
        <input placeholder="Nom complet" value={name} onChange={(e) => setName(e.target.value)} />
        <input type="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} />
        <input type="password" placeholder="Mot de passe (min. 4 caractères)" value={pwd} onChange={(e) => setPwd(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
        {error && <p className="auth-error"><AlertCircle size={13} /> {error}</p>}
        <button className="primary-btn" onClick={submit}><Plus size={15} /> Ajouter ce compte</button>
      </div>

      {saved && <p className="import-success">Compte administrateur ajouté.</p>}

      <div className="backup-hint" style={{ marginTop: 22 }}>
        <strong>Pour les enseignants</strong>
        <p>
          Les mots de passe des enseignants se gèrent depuis l'onglet <strong>Enseignants</strong> : ouvrez la fiche
          d'un enseignant pour lui assigner ou changer son mot de passe. Sans mot de passe, un enseignant ne peut
          pas se connecter.
        </p>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Nav item                                                            */
/* ------------------------------------------------------------------ */

function NavItem({ icon, label, active, onClick }) {
  return (
    <button className={`nav-item ${active ? "active" : ""}`} onClick={onClick}>
      {icon}
      <span>{label}</span>
      {active && <ChevronRight size={14} className="nav-chevron" />}
    </button>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Accueil                                                       */
/* ------------------------------------------------------------------ */

function AccueilPage({ schoolName, totalStudents, studentsInClass, setPage, setActiveClassId }) {
  const cycles = ["Maternelle", "Primaire", "Secondaire"];
  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Registre de l'établissement</p>
        <h1>{schoolName}</h1>
        <p className="lede">{totalStudents} élève{totalStudents > 1 ? "s" : ""} inscrit{totalStudents > 1 ? "s" : ""}, réparti{totalStudents > 1 ? "s" : ""} sur {CLASSES.length} classes.</p>
      </header>

      {cycles.map((cycle) => (
        <div key={cycle} className="cycle-block">
          <h3 className="cycle-title">{cycle}</h3>
          <div className="class-grid">
            {CLASSES.filter((c) => c.cycle === cycle).map((c) => {
              const n = studentsInClass(c.id).length;
              return (
                <button
                  key={c.id}
                  className="class-card"
                  onClick={() => { setActiveClassId(c.id); setPage("eleves"); }}
                >
                  <span className="class-card-name">{c.name}</span>
                  <span className="class-card-count">{n}</span>
                  <span className="class-card-label">élève{n !== 1 ? "s" : ""}</span>
                </button>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Élèves                                                        */
/* ------------------------------------------------------------------ */

function ClassPicker({ activeClassId, setActiveClassId }) {
  return (
    <select className="class-picker" value={activeClassId} onChange={(e) => setActiveClassId(e.target.value)}>
      {["Maternelle", "Primaire", "Secondaire"].map((cycle) => (
        <optgroup label={cycle} key={cycle}>
          {CLASSES.filter((c) => c.cycle === cycle).map((c) => (
            <option value={c.id} key={c.id}>{c.name}</option>
          ))}
        </optgroup>
      ))}
    </select>
  );
}

function ElevesPage({ activeClassId, setActiveClassId, students, addStudent, updateStudent, deleteStudent }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [query, setQuery] = useState("");

  const filtered = students.filter((s) =>
    `${s.firstName} ${s.lastName}`.toLowerCase().includes(query.toLowerCase())
  );

  const openNew = () => { setEditing(null); setShowForm(true); };
  const openEdit = (s) => { setEditing(s); setShowForm(true); };

  const cls = CLASSES.find((c) => c.id === activeClassId);

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Élèves</p>
        <h1>{cls.name}</h1>
        <p className="lede">{students.length} élève{students.length !== 1 ? "s" : ""} dans cette classe.</p>
      </header>

      <div className="toolbar">
        <ClassPicker activeClassId={activeClassId} setActiveClassId={setActiveClassId} />
        <div className="search-box">
          <Search size={14} />
          <input placeholder="Rechercher un élève…" value={query} onChange={(e) => setQuery(e.target.value)} />
        </div>
        <button className="primary-btn" onClick={openNew}><Plus size={15} /> Ajouter un élève</button>
      </div>

      <div className="ledger">
        <table>
          <thead>
            <tr>
              <th>Nom</th>
              <th>Prénom</th>
              <th>Naissance</th>
              <th>Sexe</th>
              <th>Parent / Tuteur</th>
              <th>Contact</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={7} className="empty-row">Aucun élève pour l'instant. Ajoutez le premier élève de la classe.</td></tr>
            )}
            {filtered.map((s) => (
              <tr key={s.id}>
                <td>{s.lastName}</td>
                <td>{s.firstName}</td>
                <td>{s.birthDate || "—"}</td>
                <td>{s.sexe || "—"}</td>
                <td>{s.parentName || "—"}</td>
                <td>{s.parentContact || "—"}</td>
                <td className="row-actions">
                  <button onClick={() => openEdit(s)} title="Modifier"><Pencil size={14} /></button>
                  <button onClick={() => deleteStudent(s.id)} title="Supprimer" className="danger"><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {showForm && (
        <StudentForm
          classId={activeClassId}
          student={editing}
          onCancel={() => setShowForm(false)}
          onSave={(data) => {
            if (editing) updateStudent(editing.id, data);
            else addStudent({ ...data, classId: activeClassId });
            setShowForm(false);
          }}
        />
      )}
    </div>
  );
}

function StudentForm({ student, onCancel, onSave }) {
  const [firstName, setFirstName] = useState(student?.firstName || "");
  const [lastName, setLastName] = useState(student?.lastName || "");
  const [birthDate, setBirthDate] = useState(student?.birthDate || "");
  const [sexe, setSexe] = useState(student?.sexe || "F");
  const [parentName, setParentName] = useState(student?.parentName || "");
  const [parentContact, setParentContact] = useState(student?.parentContact || "");

  const submit = () => {
    if (!firstName.trim() || !lastName.trim()) return;
    onSave({ firstName, lastName, birthDate, sexe, parentName, parentContact });
  };

  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h3>{student ? "Modifier l'élève" : "Nouvel élève"}</h3>
          <button type="button" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="form-grid">
          <label>Prénom<input value={firstName} onChange={(e) => setFirstName(e.target.value)} required /></label>
          <label>Nom<input value={lastName} onChange={(e) => setLastName(e.target.value)} required /></label>
          <label>Date de naissance<input type="date" value={birthDate} onChange={(e) => setBirthDate(e.target.value)} /></label>
          <label>Sexe
            <select value={sexe} onChange={(e) => setSexe(e.target.value)}>
              <option value="F">Fille</option>
              <option value="G">Garçon</option>
            </select>
          </label>
          <label>Nom du parent / tuteur<input value={parentName} onChange={(e) => setParentName(e.target.value)} /></label>
          <label>Contact (téléphone)<input value={parentContact} onChange={(e) => setParentContact(e.target.value)} /></label>
        </div>
        <div className="modal-foot">
          <button type="button" className="ghost-btn" onClick={onCancel}>Annuler</button>
          <button type="button" className="primary-btn" onClick={submit}><Save size={14} /> Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Enseignants                                                   */
/* ------------------------------------------------------------------ */

function EnseignantsPage({ teachers, addTeacher, updateTeacher, deleteTeacher }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [query, setQuery] = useState("");

  const filtered = teachers.filter((t) =>
    `${t.firstName} ${t.lastName}`.toLowerCase().includes(query.toLowerCase())
  );

  const openNew = () => { setEditing(null); setShowForm(true); };
  const openEdit = (t) => { setEditing(t); setShowForm(true); };

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Personnel enseignant</p>
        <h1>Enseignants</h1>
        <p className="lede">{teachers.length} enseignant{teachers.length !== 1 ? "s" : ""} référencé{teachers.length !== 1 ? "s" : ""}.</p>
      </header>

      <div className="toolbar">
        <div className="search-box">
          <Search size={14} />
          <input placeholder="Rechercher un enseignant…" value={query} onChange={(e) => setQuery(e.target.value)} />
        </div>
        <button className="primary-btn" onClick={openNew}><Plus size={15} /> Ajouter un enseignant</button>
      </div>

      <div className="ledger">
        <table>
          <thead>
            <tr>
              <th>Nom</th>
              <th>Prénom</th>
              <th>Matière(s)</th>
              <th>Classe(s) en charge</th>
              <th>Email</th>
              <th>Accès</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={7} className="empty-row">Aucun enseignant pour l'instant. Ajoutez le premier membre du personnel.</td></tr>
            )}
            {filtered.map((t) => (
              <tr key={t.id}>
                <td>{t.lastName}</td>
                <td>{t.firstName}</td>
                <td>{t.subjectsTaught || "—"}</td>
                <td>{t.classesInCharge || "—"}</td>
                <td>{t.email || "—"}</td>
                <td>
                  {t.password ? <span className="access-tag access-ok">Configuré</span> : <span className="access-tag access-none">Non configuré</span>}
                </td>
                <td className="row-actions">
                  <button onClick={() => openEdit(t)} title="Modifier"><Pencil size={14} /></button>
                  <button onClick={() => deleteTeacher(t.id)} title="Supprimer" className="danger"><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {showForm && (
        <TeacherForm
          teacher={editing}
          onCancel={() => setShowForm(false)}
          onSave={(data) => {
            if (editing) updateTeacher(editing.id, data);
            else addTeacher(data);
            setShowForm(false);
          }}
        />
      )}
    </div>
  );
}

function TeacherForm({ teacher, onCancel, onSave }) {
  const [firstName, setFirstName] = useState(teacher?.firstName || "");
  const [lastName, setLastName] = useState(teacher?.lastName || "");
  const [subjectsTaught, setSubjectsTaught] = useState(teacher?.subjectsTaught || "");
  const [classesInCharge, setClassesInCharge] = useState(teacher?.classesInCharge || "");
  const [phone, setPhone] = useState(teacher?.phone || "");
  const [email, setEmail] = useState(teacher?.email || "");
  const [password, setPassword] = useState("");
  const [showPwd, setShowPwd] = useState(false);
  const [error, setError] = useState("");

  const submit = () => {
    if (!firstName.trim() || !lastName.trim()) return;
    if (!teacher && !email.trim()) {
      setError("Un email est nécessaire pour que l'enseignant puisse se connecter.");
      return;
    }
    if (password && password.length < 4) {
      setError("Le mot de passe doit contenir au moins 4 caractères.");
      return;
    }
    const payload = { firstName, lastName, subjectsTaught, classesInCharge, phone, email };
    if (password) payload.password = password;
    else if (teacher?.password) payload.password = teacher.password;
    onSave(payload);
  };

  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h3>{teacher ? "Modifier l'enseignant" : "Nouvel enseignant"}</h3>
          <button type="button" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="form-grid">
          <label>Prénom<input value={firstName} onChange={(e) => setFirstName(e.target.value)} required /></label>
          <label>Nom<input value={lastName} onChange={(e) => setLastName(e.target.value)} required /></label>
          <label className="span-2">Matière(s) enseignée(s)<input value={subjectsTaught} onChange={(e) => setSubjectsTaught(e.target.value)} placeholder="ex : Français, Histoire-Géo" /></label>
          <label className="span-2">Classe(s) en charge<input value={classesInCharge} onChange={(e) => setClassesInCharge(e.target.value)} placeholder="ex : 6ème année, 7ème année" /></label>
          <label>Téléphone<input value={phone} onChange={(e) => setPhone(e.target.value)} /></label>
          <label>Email (identifiant de connexion)<input type="email" value={email} onChange={(e) => setEmail(e.target.value)} /></label>
        </div>

        <div className="teacher-access">
          <span className="auth-section-title">
            {teacher?.password ? "Changer le mot de passe" : "Assigner un mot de passe"}
          </span>
          <div className="auth-pwd-row">
            <input
              type={showPwd ? "text" : "password"}
              placeholder={teacher?.password ? "Laisser vide pour ne pas changer" : "Mot de passe (min. 4 caractères)"}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
            <button type="button" className="auth-eye" onClick={() => setShowPwd((v) => !v)}>
              {showPwd ? <EyeOff size={16} /> : <Eye size={16} />}
            </button>
          </div>
          {!teacher?.password && !password && (
            <p className="teacher-access-hint">Sans mot de passe, cet enseignant ne pourra pas se connecter à l'application.</p>
          )}
        </div>

        {error && <p className="auth-error"><AlertCircle size={13} /> {error}</p>}

        <div className="modal-foot">
          <button type="button" className="ghost-btn" onClick={onCancel}>Annuler</button>
          <button type="button" className="primary-btn" onClick={submit}><Save size={14} /> Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Emploi du temps                                               */
/* ------------------------------------------------------------------ */

function EmploiDuTempsPage({ schedules, setScheduleCell }) {
  const cycles = ["Maternelle", "Primaire", "Secondaire"];
  const [cycle, setCycle] = useState(cycles[1]);
  const cycleData = schedules[cycle] || {};

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Organisation</p>
        <h1>Emploi du temps</h1>
        <p className="lede">Un tableau hebdomadaire par cycle. Cliquez dans une case pour indiquer la matière et l'enseignant.</p>
      </header>

      <div className="toolbar">
        <div className="mode-switch">
          {cycles.map((c) => (
            <button key={c} className={cycle === c ? "active" : ""} onClick={() => setCycle(c)}>{c}</button>
          ))}
        </div>
      </div>

      <div className="ledger schedule-wrap">
        <table className="schedule-table">
          <thead>
            <tr>
              <th className="schedule-time-col">Horaire</th>
              {DAYS.map((d) => <th key={d}>{d}</th>)}
            </tr>
          </thead>
          <tbody>
            {TIME_SLOTS.map((slot) => (
              <tr key={slot.id} className={slot.pause ? "schedule-pause-row" : ""}>
                <td className="schedule-time-col mono-col">{slot.label}</td>
                {slot.pause ? (
                  <td colSpan={DAYS.length} className="schedule-pause-cell">{slot.pause}</td>
                ) : (
                  DAYS.map((d) => {
                    const key = `${slot.id}-${d}`;
                    return (
                      <td key={key}>
                        <input
                          className="schedule-input"
                          value={cycleData[key] || ""}
                          onChange={(e) => setScheduleCell(cycle, key, e.target.value)}
                          placeholder="—"
                        />
                      </td>
                    );
                  })
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Notes                                                         */
/* ------------------------------------------------------------------ */

function NotesPage({ activeClassId, setActiveClassId, activeSubjectId, setActiveSubjectId, subjects, students, grades, setGrade }) {
  const cls = CLASSES.find((c) => c.id === activeClassId);
  const subject = subjects.find((s) => s.id === activeSubjectId) || subjects[0];

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Saisie des notes</p>
        <h1>{cls.name}</h1>
        <p className="lede">Saisissez les notes sur 20 pour chaque matière. Le classement se met à jour automatiquement.</p>
      </header>

      <div className="toolbar">
        <ClassPicker activeClassId={activeClassId} setActiveClassId={setActiveClassId} />
        <select className="class-picker" value={subject?.id} onChange={(e) => setActiveSubjectId(e.target.value)}>
          {subjects.map((s) => (
            <option value={s.id} key={s.id}>{s.name} (coef. {s.coef})</option>
          ))}
        </select>
      </div>

      <div className="ledger">
        <table>
          <thead>
            <tr>
              <th>Nom</th>
              <th>Prénom</th>
              <th className="mono-col">Note / 20 — {subject?.name}</th>
            </tr>
          </thead>
          <tbody>
            {students.length === 0 && (
              <tr><td colSpan={3} className="empty-row">Aucun élève dans cette classe. Ajoutez des élèves depuis l'onglet « Élèves ».</td></tr>
            )}
            {students.map((s) => (
              <tr key={s.id}>
                <td>{s.lastName}</td>
                <td>{s.firstName}</td>
                <td className="mono-col">
                  <input
                    type="number"
                    min="0"
                    max="20"
                    step="0.25"
                    className="grade-input"
                    value={(grades[s.id] && grades[s.id][subject?.id]) ?? ""}
                    onChange={(e) => setGrade(s.id, subject.id, e.target.value)}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Classement                                                    */
/* ------------------------------------------------------------------ */

function ClassementPage({ activeClassId, setActiveClassId, students, grades, subjects }) {
  const cls = CLASSES.find((c) => c.id === activeClassId);
  const ranked = useMemo(() => rankClass(activeClassId, students, grades, subjects), [activeClassId, students, grades, subjects]);

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Classement par ordre de mérite</p>
        <h1>{cls.name}</h1>
        <p className="lede">Moyenne pondérée par coefficient, toutes matières confondues.</p>
      </header>

      <div className="toolbar">
        <ClassPicker activeClassId={activeClassId} setActiveClassId={setActiveClassId} />
      </div>

      <div className="ledger">
        <table>
          <thead>
            <tr>
              <th className="mono-col">Rang</th>
              <th>Nom</th>
              <th>Prénom</th>
              <th className="mono-col">Moyenne</th>
              <th>Mention</th>
            </tr>
          </thead>
          <tbody>
            {ranked.length === 0 && (
              <tr><td colSpan={5} className="empty-row">Aucun élève dans cette classe.</td></tr>
            )}
            {ranked.map((s) => (
              <tr key={s.id} className={s.rang === 1 ? "rank-first" : ""}>
                <td className="mono-col">
                  {s.rang === 1 ? <span className="seal">1</span> : (s.rang ?? "—")}
                </td>
                <td>{s.lastName}</td>
                <td>{s.firstName}</td>
                <td className="mono-col">{s.moyenne ?? "—"}</td>
                <td>{mention(s.moyenne)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Fiche élève (imprimable)                                      */
/* ------------------------------------------------------------------ */

function FichePage({ students, activeStudentId, setActiveStudentId, grades, subjects, schoolName }) {
  const student = students.find((s) => s.id === activeStudentId) || null;
  const cls = student ? CLASSES.find((c) => c.id === student.classId) : null;
  const ranked = student ? rankClass(student.classId, students, grades, subjects) : [];
  const rangInfo = ranked.find((s) => s.id === student?.id);
  const moyenne = student ? computeAverage(student.id, grades, subjects) : null;

  return (
    <div className="page">
      <header className="page-head no-print">
        <p className="eyebrow">Fiche élève</p>
        <h1>Relevé individuel</h1>
        <p className="lede">Sélectionnez un élève pour afficher, puis imprimer, sa fiche complète.</p>
      </header>

      <div className="toolbar no-print">
        <select className="class-picker" value={activeStudentId || ""} onChange={(e) => setActiveStudentId(e.target.value)}>
          <option value="" disabled>Choisir un élève…</option>
          {CLASSES.map((c) => {
            const list = students.filter((s) => s.classId === c.id);
            if (list.length === 0) return null;
            return (
              <optgroup label={c.name} key={c.id}>
                {list.map((s) => (
                  <option value={s.id} key={s.id}>{s.lastName} {s.firstName}</option>
                ))}
              </optgroup>
            );
          })}
        </select>
        {student && (
          <button className="primary-btn" onClick={() => window.print()}><Printer size={15} /> Imprimer la fiche</button>
        )}
      </div>

      {!student && <p className="muted no-print">Aucun élève sélectionné.</p>}

      {student && (
        <div className="fiche">
          <div className="fiche-head">
            <div>
              <p className="fiche-school">{schoolName}</p>
              <h2>Relevé de scolarité</h2>
            </div>
            <div className="seal seal-large">{cls.name.split(" ")[0]}</div>
          </div>

          <div className="fiche-info">
            <div><span>Nom</span><strong>{student.lastName}</strong></div>
            <div><span>Prénom</span><strong>{student.firstName}</strong></div>
            <div><span>Classe</span><strong>{cls.name}</strong></div>
            <div><span>Date de naissance</span><strong>{student.birthDate || "—"}</strong></div>
            <div><span>Sexe</span><strong>{student.sexe === "G" ? "Garçon" : "Fille"}</strong></div>
            <div><span>Parent / Tuteur</span><strong>{student.parentName || "—"}</strong></div>
            <div><span>Contact</span><strong>{student.parentContact || "—"}</strong></div>
          </div>

          <table className="fiche-table">
            <thead>
              <tr><th>Matière</th><th className="mono-col">Coefficient</th><th className="mono-col">Note / 20</th></tr>
            </thead>
            <tbody>
              {subjects.map((subj) => (
                <tr key={subj.id}>
                  <td>{subj.name}</td>
                  <td className="mono-col">{subj.coef}</td>
                  <td className="mono-col">{(grades[student.id] && grades[student.id][subj.id]) ?? "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className="fiche-summary">
            <div><span>Moyenne générale</span><strong>{moyenne ?? "—"} / 20</strong></div>
            <div><span>Rang dans la classe</span><strong>{rangInfo?.rang ?? "—"} / {ranked.length}</strong></div>
            <div><span>Mention</span><strong>{mention(moyenne)}</strong></div>
          </div>

          <p className="fiche-footer">Document généré automatiquement — {new Date().toLocaleDateString("fr-FR")}</p>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Discussion (chat profs / direction)                          */
/* ------------------------------------------------------------------ */

function ChatPage({ messages, sendMessage, chatName, setChatName, role }) {
  const [text, setText] = useState("");
  const [tempName, setTempName] = useState("");
  const bottomRef = useRef(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const submit = () => {
    if (!text.trim()) return;
    sendMessage(text.trim());
    setText("");
  };

  const confirmName = () => {
    if (!tempName.trim()) return;
    setChatName(tempName.trim());
  };

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Discussion</p>
        <h1>Salle des professeurs</h1>
        <p className="lede">Un espace commun pour échanger entre professeurs et direction.</p>
      </header>

      {!chatName && (
        <div className="name-prompt">
          <input
            placeholder="Votre nom (ex : Mme Diop)"
            value={tempName}
            onChange={(e) => setTempName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && confirmName()}
          />
          <button className="primary-btn" type="button" onClick={confirmName}>Rejoindre la discussion</button>
        </div>
      )}

      {chatName && (
        <>
          <div className="chat-window">
            {messages.length === 0 && <p className="muted">Aucun message pour l'instant. Lancez la discussion.</p>}
            {messages.map((m) => (
              <div className={`chat-msg ${m.role === "direction" ? "is-direction" : ""}`} key={m.id}>
                <div className="chat-msg-head">
                  <strong>{m.author}</strong>
                  <span>{m.role === "direction" ? "Direction" : "Professeur"}</span>
                  <time>{new Date(m.at).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}</time>
                </div>
                <p>{m.text}</p>
              </div>
            ))}
            <div ref={bottomRef} />
          </div>
          <div className="chat-input">
            <input
              placeholder={`Écrire en tant que ${chatName}…`}
              value={text}
              onChange={(e) => setText(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && submit()}
            />
            <button className="primary-btn" type="button" onClick={submit}><Send size={14} /> Envoyer</button>
          </div>
        </>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Direction                                                     */
/* ------------------------------------------------------------------ */

function DirectionPage({ students, studentsInClass, grades, subjects }) {
  const total = students.length;
  const maxCount = Math.max(1, ...CLASSES.map((c) => studentsInClass(c.id).length));

  const classAverages = CLASSES.map((c) => {
    const ranked = rankClass(c.id, students, grades, subjects);
    const withNote = ranked.filter((s) => s.moyenne !== null);
    const avg = withNote.length
      ? Math.round((withNote.reduce((sum, s) => sum + s.moyenne, 0) / withNote.length) * 100) / 100
      : null;
    return { ...c, avg, count: studentsInClass(c.id).length };
  });

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Vue d'ensemble — Direction</p>
        <h1>Analyse de l'établissement</h1>
        <p className="lede">{total} élèves au total, répartis sur {CLASSES.length} classes.</p>
      </header>

      <div className="stat-row">
        <div className="stat-card">
          <span className="stat-label">Effectif total</span>
          <span className="stat-value">{total}</span>
        </div>
        <div className="stat-card">
          <span className="stat-label">Classes actives</span>
          <span className="stat-value">{CLASSES.filter((c) => studentsInClass(c.id).length > 0).length} / {CLASSES.length}</span>
        </div>
        <div className="stat-card">
          <span className="stat-label">Moyenne générale école</span>
          <span className="stat-value">
            {(() => {
              const avgs = classAverages.filter((c) => c.avg !== null).map((c) => c.avg);
              return avgs.length ? (Math.round((avgs.reduce((a, b) => a + b, 0) / avgs.length) * 100) / 100) : "—";
            })()}
          </span>
        </div>
      </div>

      <h3 className="cycle-title">Effectifs par classe</h3>
      <div className="bar-chart">
        {classAverages.map((c) => (
          <div className="bar-row" key={c.id}>
            <span className="bar-label">{c.name}</span>
            <div className="bar-track">
              <div className="bar-fill" style={{ width: `${(c.count / maxCount) * 100}%` }} />
            </div>
            <span className="bar-value">{c.count}</span>
            <span className="bar-avg">{c.avg !== null ? `moy. ${c.avg}` : "—"}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Palmarès (meilleurs élèves)                                   */
/* ------------------------------------------------------------------ */

function PalmaresPage({ students, grades, subjects }) {
  const cycles = ["Maternelle", "Primaire", "Secondaire"];
  const withAvg = studentsWithAverage(students, grades, subjects);
  const ranked = withAvg.filter((s) => s.moyenne !== null).sort((a, b) => b.moyenne - a.moyenne);
  const schoolTop = ranked[0] || null;

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Palmarès</p>
        <h1>Meilleurs élèves</h1>
        <p className="lede">Le meilleur élève de chaque cycle, et de l'école toute entière.</p>
      </header>

      {schoolTop ? (
        <div className="spotlight">
          <div className="seal seal-large spotlight-seal"><Star size={22} /></div>
          <div>
            <span className="spotlight-label">Meilleur élève de l'école</span>
            <h2>{schoolTop.firstName} {schoolTop.lastName}</h2>
            <p>{schoolTop.cls?.name} — moyenne {schoolTop.moyenne} / 20</p>
          </div>
        </div>
      ) : (
        <p className="muted">Aucune note enregistrée pour l'instant.</p>
      )}

      <h3 className="cycle-title" style={{ marginTop: 28 }}>Meilleur élève par cycle</h3>
      <div className="class-grid">
        {cycles.map((cycle) => {
          const classIds = CLASSES.filter((c) => c.cycle === cycle).map((c) => c.id);
          const list = withAvg.filter((s) => classIds.includes(s.classId) && s.moyenne !== null).sort((a, b) => b.moyenne - a.moyenne);
          const top = list[0];
          return (
            <div className="podium-card" key={cycle}>
              <span className="podium-cycle">{cycle}</span>
              {top ? (
                <>
                  <strong>{top.firstName} {top.lastName}</strong>
                  <span className="podium-class">{top.cls?.name}</span>
                  <span className="podium-avg">{top.moyenne} / 20</span>
                </>
              ) : (
                <span className="muted">Pas encore de notes</span>
              )}
            </div>
          );
        })}
      </div>

      <h3 className="cycle-title" style={{ marginTop: 28 }}>Top 10 de l'école</h3>
      <div className="ledger">
        <table>
          <thead>
            <tr><th className="mono-col">Rang</th><th>Nom</th><th>Prénom</th><th>Classe</th><th className="mono-col">Moyenne</th></tr>
          </thead>
          <tbody>
            {ranked.length === 0 && <tr><td colSpan={5} className="empty-row">Aucune donnée pour l'instant.</td></tr>}
            {ranked.slice(0, 10).map((s, i) => (
              <tr key={s.id} className={i === 0 ? "rank-first" : ""}>
                <td className="mono-col">{i === 0 ? <span className="seal">1</span> : i + 1}</td>
                <td>{s.lastName}</td>
                <td>{s.firstName}</td>
                <td>{s.cls?.name}</td>
                <td className="mono-col">{s.moyenne}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Statistiques (graphiques comparatifs)                        */
/* ------------------------------------------------------------------ */

function StatistiquesPage({ students, grades, subjects, studentsInClass }) {
  const [mode, setMode] = useState("moyennes");

  const dataMoyennes = CLASSES.map((c) => {
    const ranked = rankClass(c.id, students, grades, subjects);
    const withNote = ranked.filter((s) => s.moyenne !== null);
    const avg = withNote.length ? Math.round((withNote.reduce((a, s) => a + s.moyenne, 0) / withNote.length) * 100) / 100 : 0;
    return { name: c.name, valeur: avg };
  });

  const dataEffectifs = CLASSES.map((c) => ({ name: c.name, valeur: studentsInClass(c.id).length }));

  const dataMatieres = subjects.map((subj) => {
    const notes = students
      .map((s) => grades[s.id] && grades[s.id][subj.id])
      .filter((n) => n !== undefined && n !== "" && n !== null)
      .map(Number);
    const avg = notes.length ? Math.round((notes.reduce((a, b) => a + b, 0) / notes.length) * 100) / 100 : 0;
    return { name: subj.name, valeur: avg };
  });

  const config = {
    moyennes: { data: dataMoyennes, label: "Moyenne / 20" },
    effectifs: { data: dataEffectifs, label: "Effectif" },
    matieres: { data: dataMatieres, label: "Moyenne / 20" },
  }[mode];

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Statistiques</p>
        <h1>Comparaisons</h1>
        <p className="lede">Visualisez et comparez les classes, les effectifs ou les matières.</p>
      </header>

      <div className="toolbar">
        <div className="mode-switch">
          <button className={mode === "moyennes" ? "active" : ""} onClick={() => setMode("moyennes")}>Moyennes par classe</button>
          <button className={mode === "effectifs" ? "active" : ""} onClick={() => setMode("effectifs")}>Effectifs par classe</button>
          <button className={mode === "matieres" ? "active" : ""} onClick={() => setMode("matieres")}>Moyenne par matière</button>
        </div>
      </div>

      <div className="chart-box">
        <ResponsiveContainer width="100%" height={380}>
          <BarChart data={config.data} margin={{ top: 10, right: 20, left: 0, bottom: 70 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#DCD2BC" />
            <XAxis dataKey="name" angle={-40} textAnchor="end" interval={0} tick={{ fontSize: 11, fill: "#5B6478" }} height={90} />
            <YAxis tick={{ fontSize: 11, fill: "#5B6478" }} />
            <Tooltip formatter={(v) => [v, config.label]} contentStyle={{ fontFamily: "Inter", fontSize: 12, borderRadius: 8, border: "1px solid #DCD2BC" }} />
            <Bar dataKey="valeur" fill="#2F4538" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Import (fichier Excel / CSV)                                  */
/* ------------------------------------------------------------------ */

function ImportPage({ addManyStudents }) {
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [imported, setImported] = useState(false);

  const matchClass = (raw) => {
    const n = normalizeText(raw);
    if (!n) return null;
    let found = CLASSES.find((c) => normalizeText(c.name) === n);
    if (found) return found.id;
    if (n.includes("petite")) return "ps";
    if (n.includes("moyenne")) return "ms";
    if (n.includes("grande")) return "gs";
    const numMatch = n.match(/(\d+)/);
    if (numMatch) {
      found = CLASSES.find((c) => c.id === `a${numMatch[1]}`);
      if (found) return found.id;
    }
    return null;
  };

  const getField = (row, ...keys) => {
    for (const k of keys) {
      const found = Object.keys(row).find((rk) => normalizeText(rk) === normalizeText(k));
      if (found) return row[found];
    }
    return "";
  };

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      const parsed = json
        .map((row) => {
          const rawClass = getField(row, "Classe", "Class");
          const sexeRaw = normalizeText(getField(row, "Sexe", "Genre"));
          return {
            id: uid(),
            firstName: String(getField(row, "Prénom", "Prenom", "First name") || "").trim(),
            lastName: String(getField(row, "Nom", "Last name") || "").trim(),
            birthDate: String(getField(row, "Date de naissance", "Naissance", "Date") || "").trim(),
            sexe: sexeRaw.startsWith("g") ? "G" : "F",
            parentName: String(getField(row, "Nom du parent/tuteur", "Parent", "Tuteur") || "").trim(),
            parentContact: String(getField(row, "Contact", "Téléphone", "Telephone") || "").trim(),
            classNameRaw: String(rawClass || "").trim(),
            classId: matchClass(rawClass),
          };
        })
        .filter((r) => r.firstName || r.lastName);
      setRows(parsed);
      setImported(false);
    } catch (err) {
      alert("Impossible de lire ce fichier. Vérifiez qu'il s'agit bien d'un fichier Excel (.xlsx) ou CSV.");
    }
  };

  const updateRowClass = (id, classId) => setRows((prev) => prev.map((r) => (r.id === id ? { ...r, classId } : r)));

  const confirmImport = () => {
    const valid = rows.filter((r) => r.classId);
    addManyStudents(valid.map(({ id, classNameRaw, ...rest }) => rest));
    setImported(true);
  };

  const unmatchedCount = rows.filter((r) => !r.classId).length;

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Import</p>
        <h1>Importer des élèves</h1>
        <p className="lede">Chargez un fichier Excel ou CSV : chaque élève sera automatiquement réparti dans sa classe inscrite.</p>
      </header>

      <div className="import-box">
        <label className="upload-zone">
          <Upload size={20} />
          <span>{fileName || "Choisir un fichier .xlsx, .xls ou .csv"}</span>
          <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} hidden />
        </label>
        <p className="import-hint">Colonnes reconnues : Nom, Prénom, Classe, Date de naissance, Sexe, Nom du parent/tuteur, Contact.</p>
      </div>

      {rows.length > 0 && (
        <>
          {unmatchedCount > 0 && (
            <div className="import-warning"><AlertCircle size={14} /> {unmatchedCount} élève(s) sans classe reconnue automatiquement — précisez la classe ci-dessous.</div>
          )}
          <div className="ledger">
            <table>
              <thead><tr><th>Nom</th><th>Prénom</th><th>Classe (fichier)</th><th>Classe assignée</th></tr></thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.id} className={!r.classId ? "row-warning" : ""}>
                    <td>{r.lastName}</td>
                    <td>{r.firstName}</td>
                    <td>{r.classNameRaw || "—"}</td>
                    <td>
                      <select className="cell-input" value={r.classId || ""} onChange={(e) => updateRowClass(r.id, e.target.value)}>
                        <option value="" disabled>Choisir…</option>
                        {CLASSES.map((c) => <option value={c.id} key={c.id}>{c.name}</option>)}
                      </select>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="toolbar" style={{ marginTop: 16 }}>
            <button className="primary-btn" onClick={confirmImport} disabled={rows.every((r) => !r.classId)}>
              <Upload size={15} /> Importer {rows.filter((r) => r.classId).length} élève(s)
            </button>
          </div>
          {imported && <p className="import-success">Import terminé — les élèves ont été ajoutés dans leurs classes respectives.</p>}
        </>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Sauvegarde (transfert manuel entre appareils)                 */
/* ------------------------------------------------------------------ */

function SauvegardePage({ exportBackup, importBackup, exportExcel, isOnline }) {
  const [fileName, setFileName] = useState("");
  const [confirmedFile, setConfirmedFile] = useState(null);
  const [done, setDone] = useState(false);

  const handlePick = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setConfirmedFile(file);
    setDone(false);
  };

  const runImport = () => {
    if (!confirmedFile) return;
    const ok = window.confirm(
      "Cette opération va REMPLACER toutes les données actuellement sur cet appareil par celles du fichier de sauvegarde. Continuer ?"
    );
    if (!ok) return;
    importBackup(confirmedFile);
    setDone(true);
  };

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Sauvegarde & transfert</p>
        <h1>Sauvegarde complète</h1>
        <p className="lede">
          L'application fonctionne entièrement hors-ligne : toutes les données restent sur cet appareil.
          Utilisez un fichier de sauvegarde pour transférer toutes les données (élèves, notes, enseignants,
          emploi du temps, discussion) vers un autre appareil.
        </p>
      </header>

      <div className={`conn-notice ${isOnline ? "is-online" : "is-offline"}`}>
        {isOnline
          ? "Vous êtes actuellement en ligne. L'application continuera de fonctionner normalement hors-ligne."
          : "Vous êtes actuellement hors-ligne. Pas de souci : tout continue de fonctionner et d'être enregistré sur cet appareil."}
      </div>

      <div className="backup-grid">
        <div className="backup-card">
          <Database size={20} />
          <h3>Télécharger une sauvegarde</h3>
          <p>Crée un fichier unique (.json) contenant toutes les données de l'école telles qu'elles sont sur cet appareil.</p>
          <button className="primary-btn" onClick={exportBackup}><Download size={15} /> Télécharger la sauvegarde</button>
        </div>

        <div className="backup-card">
          <Upload size={20} />
          <h3>Restaurer une sauvegarde</h3>
          <p>Charge un fichier de sauvegarde depuis un autre appareil. <strong>Remplace toutes les données actuelles.</strong></p>
          <label className="upload-zone">
            <Upload size={18} />
            <span>{fileName || "Choisir un fichier .json"}</span>
            <input type="file" accept=".json,application/json" onChange={handlePick} hidden />
          </label>
          <button className="primary-btn" onClick={runImport} disabled={!confirmedFile}>Restaurer maintenant</button>
          {done && <p className="import-success">Sauvegarde restaurée avec succès.</p>}
        </div>

        <div className="backup-card">
          <FileText size={20} />
          <h3>Export Excel (rapport)</h3>
          <p>Un fichier .xlsx en lecture seule pour consulter ou imprimer les données — ne sert pas à la restauration.</p>
          <button className="ghost-btn" onClick={exportExcel}><Download size={15} /> Exporter en Excel</button>
        </div>
      </div>

      <div className="backup-hint">
        <strong>Comment transférer les données entre le téléphone d'un professeur et l'ordinateur de la direction ?</strong>
        <p>
          Sur l'appareil source : Télécharger une sauvegarde, puis envoyez le fichier .json (WhatsApp, clé USB, email…).
          Sur l'appareil de destination : ouvrez cette page, choisissez le fichier reçu, puis Restaurer.
        </p>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Page: Matières                                                      */
/* ------------------------------------------------------------------ */

function MatieresPage({ subjects, addSubject, updateSubject, deleteSubject }) {
  const [name, setName] = useState("");
  const [coef, setCoef] = useState(1);

  const submit = () => {
    if (!name.trim()) return;
    addSubject({ name, coef: Number(coef) });
    setName("");
    setCoef(1);
  };

  return (
    <div className="page">
      <header className="page-head">
        <p className="eyebrow">Configuration</p>
        <h1>Matières & coefficients</h1>
        <p className="lede">Ces matières s'appliquent à toutes les classes et servent au calcul des moyennes.</p>
      </header>

      <div className="inline-form">
        <input placeholder="Nom de la matière" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
        <input type="number" min="1" max="10" placeholder="Coef." value={coef} onChange={(e) => setCoef(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
        <button className="primary-btn" type="button" onClick={submit}><Plus size={15} /> Ajouter</button>
      </div>

      <div className="ledger">
        <table>
          <thead><tr><th>Matière</th><th className="mono-col">Coefficient</th><th></th></tr></thead>
          <tbody>
            {subjects.map((s) => (
              <tr key={s.id}>
                <td>
                  <input className="cell-input" value={s.name} onChange={(e) => updateSubject(s.id, { name: e.target.value })} />
                </td>
                <td className="mono-col">
                  <input className="cell-input mono-col" type="number" min="1" max="10" value={s.coef} onChange={(e) => updateSubject(s.id, { coef: Number(e.target.value) })} />
                </td>
                <td className="row-actions">
                  <button onClick={() => deleteSubject(s.id)} className="danger" title="Supprimer"><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Style                                                               */
/* ------------------------------------------------------------------ */

function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');

      :root {
        --ink: #1B2A4A;
        --ink-soft: #5B6478;
        --kraft: #EDE4D3;
        --paper: #FBF8F2;
        --green: #2F4538;
        --gold: #C89A3F;
        --red: #A6453A;
        --line: #DCD2BC;
      }
      * { box-sizing: border-box; }
      .spin { animation: spin 1s linear infinite; }
      @keyframes spin { to { transform: rotate(360deg); } }

      .app {
        display: flex;
        min-height: 100vh;
        background: var(--kraft);
        color: var(--ink);
        font-family: 'Inter', sans-serif;
      }

      .sidebar {
        width: 250px;
        flex-shrink: 0;
        background: var(--ink);
        color: var(--paper);
        display: flex;
        flex-direction: column;
        padding: 20px 16px;
      }
      .brand { display: flex; align-items: center; gap: 10px; margin-bottom: 22px; }
      .brand-badge {
        width: 32px; height: 32px; border-radius: 8px;
        background: var(--gold); color: var(--ink);
        display: flex; align-items: center; justify-content: center; flex-shrink: 0;
      }
      .brand-name {
        background: transparent; border: none; color: var(--paper);
        font-family: 'Fraunces', serif; font-size: 17px; font-weight: 600;
        width: 100%; outline: none; padding: 2px 0;
      }
      .brand-name:focus { border-bottom: 1px solid var(--gold); }

      .session-badge {
        display: flex; align-items: center; justify-content: space-between; gap: 8px;
        background: rgba(255,255,255,0.08); border-radius: 8px; padding: 8px 10px; margin-bottom: 20px;
        font-size: 12px; color: rgba(255,255,255,0.85); font-weight: 600;
      }
      .session-badge button {
        background: transparent; border: none; color: rgba(255,255,255,0.6); cursor: pointer;
        display: flex; align-items: center; padding: 3px;
      }
      .session-badge button:hover { color: var(--gold); }

      .auth-screen {
        min-height: 100vh; display: flex; align-items: center; justify-content: center;
        background: var(--kraft); font-family: 'Inter', sans-serif; padding: 20px;
      }
      .auth-card {
        background: var(--paper); border: 1px solid var(--line); border-radius: 14px;
        padding: 32px; width: 100%; max-width: 380px; display: flex; flex-direction: column; gap: 6px;
      }
      .auth-badge { width: 40px; height: 40px; border-radius: 10px; margin-bottom: 10px; }
      .auth-card h1 { font-family: 'Fraunces', serif; font-size: 24px; margin: 0 0 6px; color: var(--ink); }
      .auth-card .lede { margin-bottom: 14px; }
      .auth-section { display: flex; flex-direction: column; gap: 8px; margin-bottom: 14px; }
      .auth-section-title { font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-soft); font-weight: 700; }
      .auth-card input {
        border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px; font-size: 13.5px;
        font-family: 'Inter', sans-serif; width: 100%; color: var(--ink);
      }
      .auth-tabs { margin-bottom: 16px; }
      .auth-email { border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px; font-size: 13.5px; font-family: 'Inter', sans-serif; width: 100%; margin-bottom: 8px; color: var(--ink); }
      .auth-note { font-size: 11.5px; color: var(--ink-soft); margin-top: 14px; text-align: center; line-height: 1.4; }
      .you-tag { font-size: 10.5px; color: var(--ink-soft); font-weight: 400; }
      .access-tag { font-size: 11px; font-weight: 600; padding: 3px 8px; border-radius: 20px; }
      .access-tag.access-ok { background: rgba(47,69,56,0.12); color: var(--green); }
      .access-tag.access-none { background: rgba(166,69,58,0.1); color: var(--red); }
      .teacher-access { border-top: 1px dashed var(--line); margin-top: 16px; padding-top: 14px; display: flex; flex-direction: column; gap: 8px; }
      .teacher-access-hint { font-size: 11.5px; color: var(--ink-soft); margin: 0; }
      .auth-pwd-row { display: flex; align-items: center; gap: 8px; border: 1px solid var(--line); border-radius: 8px; padding: 2px 6px 2px 2px; margin-bottom: 4px; }
      .auth-pwd-row input { border: none; }
      .auth-eye { background: transparent; border: none; color: var(--ink-soft); cursor: pointer; display: flex; padding: 6px; }
      .auth-error { display: flex; align-items: center; gap: 6px; color: var(--red); font-size: 12.5px; margin: 6px 0 2px; }
      .auth-submit { width: 100%; justify-content: center; margin-top: 12px; padding: 11px; }

      .nav { display: flex; flex-direction: column; gap: 2px; flex: 1; }
      .nav-item {
        display: flex; align-items: center; gap: 10px; background: transparent; border: none;
        color: rgba(255,255,255,0.75); padding: 9px 10px; border-radius: 7px; cursor: pointer;
        font-size: 13.5px; font-family: 'Inter', sans-serif; text-align: left;
      }
      .nav-item:hover { background: rgba(255,255,255,0.06); color: var(--paper); }
      .nav-item.active { background: var(--green); color: var(--paper); font-weight: 600; }
      .nav-item span { flex: 1; }
      .nav-chevron { opacity: 0.6; }

      .sidebar-foot { margin-top: auto; padding-top: 14px; }
      .export-btn {
        width: 100%; display: flex; align-items: center; justify-content: center; gap: 8px;
        background: var(--gold); color: var(--ink); border: none; border-radius: 8px;
        padding: 10px; font-weight: 700; font-size: 13px; cursor: pointer; font-family: 'Inter', sans-serif;
      }
      .save-state {
        margin-top: 8px; font-size: 11px; color: rgba(255,255,255,0.5);
        display: flex; align-items: center; gap: 4px; justify-content: center; height: 14px;
      }
      .conn-badge {
        display: flex; align-items: center; justify-content: center; gap: 5px;
        font-size: 10.5px; padding: 6px 8px; border-radius: 6px; margin-bottom: 10px;
        text-align: center; line-height: 1.3;
      }
      .conn-badge.is-online { background: rgba(255,255,255,0.06); color: rgba(255,255,255,0.55); }
      .conn-badge.is-offline { background: rgba(200,154,63,0.18); color: var(--gold); font-weight: 600; }

      /* Sauvegarde */
      .conn-notice { border-radius: 8px; padding: 12px 16px; font-size: 12.5px; margin-bottom: 20px; }
      .conn-notice.is-online { background: var(--paper); border: 1px solid var(--line); color: var(--ink-soft); }
      .conn-notice.is-offline { background: rgba(200,154,63,0.12); border: 1px solid var(--gold); color: #7a5b1e; }
      .backup-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-bottom: 20px; }
      .backup-card { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 18px; display: flex; flex-direction: column; gap: 8px; color: var(--green); }
      .backup-card h3 { font-family: 'Fraunces', serif; font-size: 15px; margin: 2px 0 0; color: var(--ink); }
      .backup-card p { font-size: 12px; color: var(--ink-soft); margin: 0 0 6px; line-height: 1.5; }
      .backup-hint { background: var(--kraft); border-radius: 10px; padding: 16px 18px; font-size: 12.5px; color: var(--ink-soft); }
      .backup-hint strong { display: block; color: var(--ink); font-size: 13px; margin-bottom: 4px; }

      .main { flex: 1; padding: 40px 48px; overflow-x: auto; }
      .page { max-width: 980px; }

      .page-head { margin-bottom: 26px; }
      .eyebrow { font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.08em; color: var(--ink-soft); font-weight: 600; margin: 0 0 6px; }
      .page-head h1 { font-family: 'Fraunces', serif; font-size: 32px; font-weight: 600; margin: 0 0 8px; }
      .lede { color: var(--ink-soft); font-size: 14px; margin: 0; }

      .cycle-block { margin-bottom: 28px; }
      .cycle-title { font-family: 'Fraunces', serif; font-size: 16px; font-weight: 600; margin: 0 0 12px; color: var(--green); }
      .class-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 10px; }
      .class-card {
        background: var(--paper); border: 1px solid var(--line); border-radius: 8px; padding: 14px;
        text-align: left; cursor: pointer; display: flex; flex-direction: column; gap: 2px;
      }
      .class-card:hover { border-color: var(--green); }
      .class-card-name { font-size: 12.5px; font-weight: 600; color: var(--ink-soft); }
      .class-card-count { font-family: 'IBM Plex Mono', monospace; font-size: 24px; font-weight: 600; color: var(--ink); }
      .class-card-label { font-size: 11px; color: var(--ink-soft); }

      .toolbar { display: flex; align-items: center; gap: 10px; margin-bottom: 18px; flex-wrap: wrap; }
      .class-picker {
        background: var(--paper); border: 1px solid var(--line); border-radius: 7px; padding: 9px 12px;
        font-size: 13px; color: var(--ink); font-family: 'Inter', sans-serif;
      }
      .search-box {
        display: flex; align-items: center; gap: 6px; background: var(--paper); border: 1px solid var(--line);
        border-radius: 7px; padding: 8px 12px; flex: 1; min-width: 160px; color: var(--ink-soft);
      }
      .search-box input { border: none; outline: none; background: transparent; font-size: 13px; flex: 1; color: var(--ink); }

      .primary-btn {
        display: flex; align-items: center; gap: 6px; background: var(--green); color: var(--paper);
        border: none; border-radius: 7px; padding: 9px 14px; font-size: 13px; font-weight: 600; cursor: pointer;
        font-family: 'Inter', sans-serif; white-space: nowrap;
      }
      .ghost-btn {
        background: transparent; border: 1px solid var(--line); color: var(--ink-soft); border-radius: 7px;
        padding: 9px 14px; font-size: 13px; cursor: pointer; font-family: 'Inter', sans-serif;
      }

      .ledger { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; overflow: hidden; }
      table { width: 100%; border-collapse: collapse; }
      th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--ink-soft); padding: 12px 16px; border-bottom: 1px solid var(--line); font-weight: 600; }
      td { padding: 11px 16px; border-bottom: 1px solid var(--line); font-size: 13.5px; }
      tr:last-child td { border-bottom: none; }
      .mono-col { font-family: 'IBM Plex Mono', monospace; }
      .empty-row { text-align: center; color: var(--ink-soft); padding: 30px; font-style: italic; }
      .row-actions { display: flex; gap: 8px; }
      .row-actions button { background: transparent; border: none; color: var(--ink-soft); cursor: pointer; padding: 4px; }
      .row-actions button.danger:hover { color: var(--red); }
      .row-actions button:hover { color: var(--ink); }

      .grade-input {
        width: 68px; border: 1px solid var(--line); border-radius: 6px; padding: 6px 8px;
        font-family: 'IBM Plex Mono', monospace; font-size: 13px; text-align: center;
      }
      .cell-input { border: 1px solid transparent; background: transparent; font-size: 13.5px; padding: 4px 6px; border-radius: 5px; width: 100%; font-family: 'Inter', sans-serif; }
      .cell-input:focus { border-color: var(--line); background: var(--kraft); outline: none; }
      .cell-input.mono-col { font-family: 'IBM Plex Mono', monospace; width: 60px; }

      .rank-first { background: rgba(200, 154, 63, 0.08); }
      .seal {
        display: inline-flex; align-items: center; justify-content: center;
        width: 26px; height: 26px; border-radius: 50%; border: 2px solid var(--gold);
        color: var(--gold); font-weight: 700; font-family: 'IBM Plex Mono', monospace; font-size: 13px;
      }

      .inline-form { display: flex; gap: 10px; margin-bottom: 18px; }
      .inline-form input { border: 1px solid var(--line); border-radius: 7px; padding: 9px 12px; font-size: 13px; font-family: 'Inter', sans-serif; }
      .inline-form input:first-child { flex: 1; }
      .inline-form input:nth-child(2) { width: 90px; }

      .stat-row { display: flex; gap: 14px; margin-bottom: 28px; }
      .stat-card { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 18px 20px; flex: 1; }
      .stat-label { display: block; font-size: 11.5px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 8px; font-weight: 600; }
      .stat-value { font-family: 'Fraunces', serif; font-size: 30px; font-weight: 600; color: var(--ink); }

      .bar-chart { display: flex; flex-direction: column; gap: 8px; }
      .bar-row { display: grid; grid-template-columns: 130px 1fr 34px 90px; align-items: center; gap: 12px; }
      .bar-label { font-size: 12.5px; color: var(--ink-soft); }
      .bar-track { background: var(--paper); border: 1px solid var(--line); border-radius: 5px; height: 16px; overflow: hidden; }
      .bar-fill { background: var(--green); height: 100%; border-radius: 5px 0 0 5px; }
      .bar-value { font-family: 'IBM Plex Mono', monospace; font-size: 12.5px; text-align: right; }
      .bar-avg { font-size: 11.5px; color: var(--ink-soft); }

      .modal-backdrop {
        position: fixed; inset: 0; background: rgba(27,42,74,0.45); display: flex;
        align-items: center; justify-content: center; z-index: 50; padding: 20px;
      }
      .modal { background: var(--paper); border-radius: 12px; padding: 22px; width: 100%; max-width: 460px; }
      .modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
      .modal-head h3 { font-family: 'Fraunces', serif; font-size: 18px; margin: 0; }
      .modal-head button { background: transparent; border: none; cursor: pointer; color: var(--ink-soft); }
      .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 18px; }
      .form-grid label { display: flex; flex-direction: column; gap: 5px; font-size: 12px; color: var(--ink-soft); font-weight: 600; }
      .form-grid input, .form-grid select {
        border: 1px solid var(--line); border-radius: 7px; padding: 8px 10px; font-size: 13px; font-family: 'Inter', sans-serif; color: var(--ink);
      }
      .modal-foot { display: flex; justify-content: flex-end; gap: 10px; }

      .muted { color: var(--ink-soft); font-style: italic; }

      /* Fiche imprimable */
      .fiche { background: var(--paper); border: 1px solid var(--line); border-radius: 12px; padding: 32px; max-width: 640px; }
      .fiche-head { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid var(--ink); padding-bottom: 16px; margin-bottom: 20px; }
      .fiche-school { font-size: 12px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.06em; margin: 0 0 4px; }
      .fiche-head h2 { font-family: 'Fraunces', serif; font-size: 24px; margin: 0; }
      .seal-large {
        width: 64px; height: 64px; font-size: 12px; border-width: 2px; transform: rotate(-6deg);
        text-align: center; line-height: 1.1;
      }
      .fiche-info { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 24px; margin-bottom: 22px; }
      .fiche-info div { display: flex; justify-content: space-between; border-bottom: 1px dotted var(--line); padding-bottom: 4px; }
      .fiche-info span { font-size: 12px; color: var(--ink-soft); }
      .fiche-info strong { font-size: 13px; }
      .fiche-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
      .fiche-table th { font-size: 11px; text-transform: uppercase; color: var(--ink-soft); text-align: left; padding: 8px 4px; border-bottom: 1px solid var(--ink); }
      .fiche-table td { padding: 8px 4px; border-bottom: 1px solid var(--line); font-size: 13px; }
      .fiche-summary { display: flex; gap: 20px; background: var(--kraft); border-radius: 8px; padding: 14px 18px; margin-bottom: 14px; }
      .fiche-summary div { display: flex; flex-direction: column; gap: 2px; }
      .fiche-summary span { font-size: 11px; color: var(--ink-soft); text-transform: uppercase; }
      .fiche-summary strong { font-family: 'Fraunces', serif; font-size: 18px; }
      .fiche-footer { font-size: 10.5px; color: var(--ink-soft); text-align: right; margin: 0; }

      .form-grid label.span-2 { grid-column: span 2; }

      /* Emploi du temps */
      .schedule-wrap { overflow-x: auto; }
      .schedule-table { min-width: 780px; }
      .schedule-time-col { white-space: nowrap; width: 130px; }
      .schedule-pause-row td { background: var(--kraft); }
      .schedule-pause-cell { text-align: center; font-size: 12px; color: var(--ink-soft); font-style: italic; }
      .schedule-input {
        width: 100%; border: 1px solid transparent; background: transparent; border-radius: 6px;
        padding: 6px 8px; font-size: 12.5px; font-family: 'Inter', sans-serif; color: var(--ink);
      }
      .schedule-input:hover { border-color: var(--line); }
      .schedule-input:focus { border-color: var(--green); background: var(--paper); outline: none; }

      /* Discussion */
      .chat-window { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 16px; height: 420px; overflow-y: auto; display: flex; flex-direction: column; gap: 10px; margin-bottom: 12px; }
      .chat-msg { max-width: 68%; background: var(--kraft); border-radius: 10px; padding: 8px 12px; align-self: flex-start; }
      .chat-msg.is-direction { align-self: flex-end; background: rgba(47,69,56,0.12); }
      .chat-msg-head { display: flex; gap: 8px; align-items: baseline; margin-bottom: 2px; }
      .chat-msg-head strong { font-size: 12.5px; }
      .chat-msg-head span { font-size: 10px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.04em; }
      .chat-msg-head time { font-size: 10px; color: var(--ink-soft); margin-left: auto; }
      .chat-msg p { margin: 0; font-size: 13.5px; }
      .chat-input { display: flex; gap: 10px; }
      .chat-input input { flex: 1; border: 1px solid var(--line); border-radius: 8px; padding: 10px 14px; font-size: 13.5px; font-family: 'Inter', sans-serif; }
      .name-prompt { display: flex; gap: 10px; max-width: 380px; }
      .name-prompt input { flex: 1; border: 1px solid var(--line); border-radius: 8px; padding: 9px 12px; font-size: 13px; font-family: 'Inter', sans-serif; }

      /* Palmarès */
      .spotlight { display: flex; align-items: center; gap: 20px; background: var(--ink); color: var(--paper); border-radius: 12px; padding: 22px 26px; }
      .spotlight-seal { background: transparent; border-color: var(--gold); color: var(--gold); }
      .spotlight-label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: rgba(255,255,255,0.6); }
      .spotlight h2 { font-family: 'Fraunces', serif; font-size: 22px; margin: 4px 0; }
      .spotlight p { margin: 0; color: rgba(255,255,255,0.75); font-size: 13px; }
      .podium-card { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 16px; display: flex; flex-direction: column; gap: 4px; }
      .podium-cycle { font-size: 11px; text-transform: uppercase; color: var(--ink-soft); font-weight: 600; letter-spacing: 0.05em; }
      .podium-card strong { font-size: 15px; }
      .podium-class { font-size: 12px; color: var(--ink-soft); }
      .podium-avg { font-family: 'IBM Plex Mono', monospace; font-weight: 600; color: var(--green); }

      /* Statistiques */
      .mode-switch { display: flex; background: var(--paper); border: 1px solid var(--line); border-radius: 8px; padding: 3px; flex-wrap: wrap; }
      .mode-switch button { background: transparent; border: none; color: var(--ink-soft); font-size: 12.5px; font-weight: 600; padding: 8px 12px; border-radius: 6px; cursor: pointer; font-family: 'Inter', sans-serif; white-space: nowrap; }
      .mode-switch button.active { background: var(--green); color: var(--paper); }
      .chart-box { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 16px 8px 6px; }

      /* Import */
      .import-box { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 20px; margin-bottom: 18px; }
      .upload-zone { display: flex; align-items: center; gap: 10px; border: 1.5px dashed var(--line); border-radius: 8px; padding: 16px; cursor: pointer; color: var(--ink-soft); font-size: 13px; }
      .upload-zone:hover { border-color: var(--green); color: var(--ink); }
      .import-hint { font-size: 11.5px; color: var(--ink-soft); margin: 10px 0 0; }
      .import-warning { display: flex; align-items: center; gap: 6px; background: rgba(166,69,58,0.1); color: var(--red); border-radius: 8px; padding: 10px 14px; font-size: 12.5px; margin-bottom: 14px; }
      .row-warning td { background: rgba(166,69,58,0.05); }
      .import-success { margin-top: 12px; color: var(--green); font-size: 13px; font-weight: 600; }

      @media print {
        .sidebar, .no-print { display: none !important; }
        .main { padding: 0; }
        .app { background: white; }
        .fiche { border: none; box-shadow: none; }
      }
    `}</style>
  );
}
