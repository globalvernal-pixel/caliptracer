import React, { useState, useEffect } from 'react';
import LoginPage from './LoginPage.jsx';
import * as XLSX from 'xlsx';
import { jsPDF } from "jspdf";
import autoTable from 'jspdf-autotable';
import html2canvas from 'html2canvas';
import logo from './logo.png';
import { 
  Sparkles, 
  Layers, 
  Users, 
  Search, 
  PlusCircle, 
  Sliders, 
  Star, 
  Hash, 
  Plus, 
  Minus, 
  Check, 
  Trash2,
  X,
  Calendar,
  ChevronLeft,
  FileText,
  Share2,
  Copy,
  Send,
  Table,
  Lock,
  Download,
  GraduationCap,
  School,
  BookOpen,
  Award,
  ChevronRight,
  UserCheck,
  TrendingUp,
  BarChart3,
  ShieldCheck,
  CheckCircle2,
  MessageSquare,
  AlertCircle,
  AlertTriangle,
  Settings,
  Eye,
  EyeOff,
  ArrowUp,
  ArrowDown,
  Pencil,
  Sun,
  List,
  Camera,
  Smartphone
} from 'lucide-react';

// Initial seed data for the 8 classes with demo student names and score records
const INITIAL_STUDENTS = [];

// Initial seed data for Mentors
const INITIAL_MENTORS = [];

const PREDEFINED_SPOT_FINES = [
  { reason: 'MISSING NOTE BOOK', amount: 200 },
  { reason: 'MISSING CALLING CARD', amount: 300 },
  { reason: 'MISSING HALL TICKET', amount: 300 },
  { reason: 'MISSING DIARY', amount: 500 },
  { reason: 'MISSING T BOOK', amount: 500 },
  { reason: 'MISSING ENTIRE ID CARD', amount: 500 },
  { reason: 'MISSING K BOOK', amount: 1000 },
  { reason: 'MISSING PASSPORT', amount: 2000 }
];

const PREDEFINED_NEAT_REASONS = [
  { reason: 'Socks', count: 5 },
  { reason: 'ID card', count: 5 },
  { reason: 'Belt', count: 5 },
  { reason: 'Shirt Tucked out', count: 5 },
  { reason: 'Different Color Scarf', count: 5 },
  { reason: 'Wearing informal shoe', count: 5 },
  { reason: 'Different Color Shoe and Socks', count: 5 },
  { reason: 'Pant', count: 10 },
  { reason: 'Shirt', count: 10 },
  { reason: 'Shoe', count: 10 }
];

const ROOM_STUDENT_MAPPING = {
  '19': [
    { id: 'r19-1', name: 'MUHAMMED UNAIS', class: 'S1B' },
    { id: 'r19-2', name: 'MUHAMMED SHAMMAS', class: 'C1B' },
    { id: 'r19-3', name: 'MUHAMMED HADIN FARHAN', class: 'C1B' },
    { id: 'r19-4', name: 'AHAMMAD SHADHIL', class: 'C1B' },
    { id: 'r19-5', name: 'IJLAN THOTUNGAL', class: 'C1B' },
    { id: 'r19-6', name: 'MUHAMMED HAMMAZ', class: 'C1B' }
  ],
  '20': [
    { id: 'r20-1', name: 'ABDULLAMUZAMMIL', class: 'S1B' },
    { id: 'r20-2', name: 'ABDULHADHI S.S', class: 'S1B' },
    { id: 'r20-3', name: 'MUHAMMED NIHAL B', class: 'C1B' },
    { id: 'r20-4', name: 'ABDUL SAMAD CA', class: 'C1C' },
    { id: 'r20-5', name: 'MOHAMMED MUJTHABA', class: 'C1C' },
    { id: 'r20-6', name: 'ZAINUL ABIDHEEN', class: 'C1C' }
  ],
  '21': [
    { id: 'r21-1', name: 'REYHAN ANEES', class: 'S2B' },
    { id: 'r21-2', name: 'SHAHIR O', class: 'S2B' },
    { id: 'r21-3', name: 'DILHAQ HASSAN', class: 'S2B' },
    { id: 'r21-4', name: 'MUHAMMED IMRAN', class: 'S2B' },
    { id: 'r21-5', name: 'MUHAMMED P', class: 'C2B' },
    { id: 'r21-6', name: 'MUHAMMED FARHAN', class: 'C2B' },
    { id: 'r21-7', name: 'MUHAMMED FADHIL RN', class: 'C2B' },
    { id: 'r21-8', name: 'MUHAMMED YOUSUF', class: 'C2B' }
  ],
  '22': [
    { id: 'r22-1', name: 'MUHAMMED MUJTHABA', class: 'S1B' },
    { id: 'r22-2', name: 'HASHIN EH', class: 'C1B' },
    { id: 'r22-3', name: 'FAIHAN FIROZ', class: 'S1B' },
    { id: 'r22-4', name: 'MUHAMMED FAAZ KP', class: 'C1B' },
    { id: 'r22-5', name: 'ALI MIYAN', class: 'C1B' },
    { id: 'r22-6', name: 'MUHAMMED HASHIM', class: 'C1C' },
    { id: 'r22-7', name: 'SHAMMAS ASLAM', class: 'C1B' }
  ],
  '124': [
    { id: 'r124-1', name: 'MINHAJ', class: 'S2B' },
    { id: 'r124-2', name: 'SHAHAN SHABEEB', class: 'S2B' },
    { id: 'r124-3', name: 'MAHDI MUHAMMED', class: 'S2B' },
    { id: 'r124-4', name: 'MUHAMMED SAHAL T', class: 'S2B' },
    { id: 'r124-5', name: 'MUFTI MUHAMMED', class: 'C2B' },
    { id: 'r124-6', name: 'AMAAN MOHAMMED RAFHEEK', class: 'C2B' }
  ],
  '125': [
    { id: 'r125-1', name: 'MUHAMMED NIHAL PB', class: 'S1B' },
    { id: 'r125-2', name: 'MOHAMMED FADI KT', class: 'S1B' },
    { id: 'r125-3', name: 'SAVAD M', class: 'C1B' },
    { id: 'r125-4', name: 'ABDUL HADHI K.S', class: 'C1B' },
    { id: 'r125-5', name: 'MUHAMMED HANAN MATTAI', class: 'C1B' },
    { id: 'r125-6', name: 'FAVAS M', class: 'C1C' }
  ],
  '126': [
    { id: 'r126-1', name: 'MOHAMMED NIHAL KP', class: 'S1B' },
    { id: 'r126-2', name: 'MUHAMMED SHEZIN SADIKH', class: 'S1B' },
    { id: 'r126-3', name: 'AYAZ MUHAMMED', class: 'C1B' },
    { id: 'r126-4', name: 'MUHAMMED MUNAZIR', class: 'C1C' },
    { id: 'r126-5', name: 'ADHIL HANAN', class: 'C1C' },
    { id: 'r126-6', name: 'SAHL ABDULLA', class: 'C1C' }
  ],
  '127': [
    { id: 'r127-1', name: 'RIZAN MUHAMMED', class: 'S1B' },
    { id: 'r127-2', name: 'SHABIL ASHRAF M K', class: 'S1B' },
    { id: 'r127-3', name: 'MUHAMMED ATHIF', class: 'C1C' },
    { id: 'r127-4', name: 'MUHAMMED YOUNUS', class: 'C1C' },
    { id: 'r127-5', name: 'MUHAMMED AMEEN ASHFAQ', class: 'C1C' },
    { id: 'r127-6', name: 'SHAZIN NOUSHAD', class: 'C1B' }
  ],
  '128': [
    { id: 'r128-1', name: 'AYAAN MOHAMMAD', class: 'S1B' },
    { id: 'r128-2', name: 'AHMED SAHL SAJID', class: 'S1B' },
    { id: 'r128-3', name: 'MUHAMMED MISHAL', class: 'C1C' },
    { id: 'r128-4', name: 'RAMIN MUHAMMED', class: 'C1C' },
    { id: 'r128-5', name: 'MUHAMMED NADUVILOTHI', class: 'C1C' },
    { id: 'r128-6', name: 'SHEZIN HAMDAN', class: 'C1B' }
  ],
  '108': [
    { id: 'r108-1', name: 'MUHAMMED NIHAD K', class: 'S1B' },
    { id: 'r108-2', name: 'AHLAN K', class: 'S1B' },
    { id: 'r108-3', name: 'MUHAMMED ABAAN', class: 'C1C' },
    { id: 'r108-4', name: 'MUADDAB KODAMBI', class: 'C1B' },
    { id: 'r108-5', name: 'HUMAYL MUHAMMED', class: 'C1B' },
    { id: 'r108-6', name: 'ABDULLAH WAIZ.N.P', class: 'C1B' }
  ],
  '109': [
    { id: 'r109-1', name: 'SAHIL MOHAMMED SAHEER', class: 'S1B' },
    { id: 'r109-2', name: 'HADI MUHAMMAD', class: 'S1B' },
    { id: 'r109-3', name: 'AMAN ABDULLA KT', class: 'C1B' },
    { id: 'r109-4', name: 'MUHAMMED FADHIL', class: 'C1B' },
    { id: 'r109-5', name: 'MOHAMMED', class: 'C1C' },
    { id: 'r109-6', name: 'RAFAN MUHAMMED', class: 'C1C' }
  ],
  '110': [
    { id: 'r110-1', name: 'HAANI MUAAD', class: 'S1B' },
    { id: 'r110-2', name: 'AMAN USMAN', class: 'S1B' },
    { id: 'r110-3', name: 'JAZEEM HANEEFA', class: 'C1B' },
    { id: 'r110-4', name: 'AHMAD MUNEER', class: 'C1B' },
    { id: 'r110-5', name: 'FAID SANEEN', class: 'C1B' },
    { id: 'r110-6', name: 'REHAN MUHAMMED SAHIR', class: 'C1C' }
  ],
  '215': [
    { id: 'r215-1', name: 'HAROON RASHEED', class: 'S1B' },
    { id: 'r215-2', name: 'HADI MUHAMMED E P', class: 'S1B' },
    { id: 'r215-3', name: 'AZMI MUHAMMED ASLAM', class: 'C1C' },
    { id: 'r215-4', name: 'RADIN T', class: 'C1C' },
    { id: 'r215-5', name: 'FADHIL ANWAR BASHEER', class: 'C1C' },
    { id: 'r215-6', name: 'SHAZIN HUSSAIN', class: 'C1C' }
  ],
  '216': [
    { id: 'r216-1', name: 'MOHAMMED AJMAL', class: 'S1B' },
    { id: 'r216-2', name: 'AHAMMED MAJID', class: 'S1B' },
    { id: 'r216-3', name: 'FAAZ MUHAMMED PS', class: 'C1B' },
    { id: 'r216-4', name: 'ZUNNOON P', class: 'C1C' },
    { id: 'r216-5', name: 'MUHAMMED RASIL MP', class: 'C1C' },
    { id: 'r216-6', name: 'MUHAMMED AMAN TP', class: 'C1B' }
  ],
  '217': [
    { id: 'r217-1', name: 'MUHAMMED HAMID TK', class: 'S2B' },
    { id: 'r217-2', name: 'MUHAMMED SWALAH', class: 'S2B' },
    { id: 'r217-3', name: 'RAZI MUHAMMED VF', class: 'C2B' },
    { id: 'r217-4', name: 'MUHAMMED MINHAI PC', class: 'C2C' },
    { id: 'r217-5', name: 'MUHAMMED RASHID T', class: 'C2B' },
    { id: 'r217-6', name: 'MUHAMMED MUHSIN K', class: 'C2B' }
  ],
  '231': [
    { id: 'r231-1', name: 'MUHAMMED ATHIQ ANZAD', class: 'S1B' },
    { id: 'r231-2', name: 'MUHAMMED ESHAN P', class: 'S1B' },
    { id: 'r231-3', name: 'AYAZ JABIR', class: 'C1B' },
    { id: 'r231-4', name: 'LUQMANUL HAKEEM', class: 'C1C' },
    { id: 'r231-5', name: 'MUHAMMED YASEEN', class: 'C1C' },
    { id: 'r231-6', name: 'MUHAMMMED FAAZ', class: 'C1C' }
  ],
  '232': [
    { id: 'r232-1', name: 'ALAN BACKER KP', class: 'C1B' },
    { id: 'r232-2', name: 'MUHAMMED YASEEN V', class: 'S1B' },
    { id: 'r232-3', name: 'AMIN RIHAN', class: 'C1C' },
    { id: 'r232-4', name: 'AZAAN MAHWISH', class: 'C1C' },
    { id: 'r232-5', name: 'MUHAMMED RASEEM', class: 'C1B' },
    { id: 'r232-6', name: 'MUHAMMED SIDHAN VP', class: 'C1B' }
  ],
  '233': [
    { id: 'r233-1', name: 'AMEEN PV', class: 'S2B' },
    { id: 'r233-2', name: 'MUHAMMED HANAN ANSAR', class: 'S2B' },
    { id: 'r233-3', name: 'MUHAMMMED SHAFEEQ', class: 'S2B' },
    { id: 'r233-4', name: 'HATHIM ZAMAN', class: 'S2B' },
    { id: 'r233-5', name: 'ALI SHAMIL', class: 'S2B' },
    { id: 'r233-6', name: 'AHAMMED YASEEN BIN SHABEER', class: 'C2B' },
    { id: 'r233-7', name: 'MUHAMMED CP', class: 'C2B' },
    { id: 'r233-8', name: 'MUHAMMED SINAN KT', class: 'C2B' }
  ],
  '234': [
    { id: 'r234-1', name: 'RAZIN MUHAMMED', class: 'S1B' },
    { id: 'r234-2', name: 'RISWAN M', class: 'S1B' },
    { id: 'r234-3', name: 'SIDAN HASSAN', class: 'C1C' },
    { id: 'r234-4', name: 'MUHAMMED YASEEN MP', class: 'C1C' },
    { id: 'r234-5', name: 'HANIN AHMED FIZIN', class: 'C1C' },
    { id: 'r234-6', name: 'MUHAMMED HAZEEN', class: 'C1C' }
  ],
  '235': [
    { id: 'r235-1', name: 'MUHAMMED AFIL', class: 'S2B' },
    { id: 'r235-2', name: 'IHAN INTHIKAF', class: 'S2B' },
    { id: 'r235-3', name: 'MUHAMMED IZAN DARVISH', class: 'S2B' },
    { id: 'r235-4', name: 'IHSAN ABDULLA', class: 'S2B' },
    { id: 'r235-5', name: 'LABEEB NOUSHAD', class: 'S2B' },
    { id: 'r235-6', name: 'ATHIF THUFAIL N', class: 'C2C' },
    { id: 'r235-7', name: 'HAROON ASLAM', class: 'C2C' },
    { id: 'r235-8', name: 'ZAHI ABOOBACKER VENGAT', class: 'C2C' }
  ]
};

// CLASSES array is now managed via state

const parseBulkText = (text, allStudents, currentClass) => {
  const regex = /\b(\d+)\s*(stars?|tallies|tally)\b/gi;
  let matches = [];
  let match;
  while ((match = regex.exec(text)) !== null) {
    matches.push({
      index: match.index,
      value: match[0],
      amount: parseInt(match[1], 10),
      type: match[2].toLowerCase().startsWith('star') ? 'star' : 'tally'
    });
  }

  if (matches.length === 0) return [];

  const results = [];
  
  for (let i = 0; i < matches.length; i++) {
    const currentMatch = matches[i];
    const startIndex = i === 0 ? 0 : (matches[i - 1].index + matches[i - 1].value.length);
    let chunk = text.substring(startIndex, currentMatch.index).trim();
    
    let matchedStudent = null;
    let matchType = '';
    let reason = '';
    
    // Replace non-alphanumeric (excluding space) with spaces and collapse spaces
    const normalizedChunk = chunk.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
    
    const searchPool = [
      ...allStudents.filter(s => s.class === currentClass),
      ...allStudents.filter(s => s.class !== currentClass)
    ];
    
    // 1. Try to match full name as suffix of the chunk
    for (const student of searchPool) {
      const studentNameNorm = student.name.toLowerCase().trim();
      if (normalizedChunk.endsWith(studentNameNorm)) {
        matchedStudent = student;
        matchType = 'full';
        const nameIdx = chunk.toLowerCase().lastIndexOf(student.name.toLowerCase());
        reason = chunk.substring(0, nameIdx).trim();
        break;
      }
    }
    
    // 2. Try to match first name as suffix of the chunk
    if (!matchedStudent) {
      for (const student of searchPool) {
        const firstName = student.name.split(' ')[0].toLowerCase().trim();
        if (firstName.length > 1 && normalizedChunk.endsWith(firstName)) {
          matchedStudent = student;
          matchType = 'first';
          const nameIdx = chunk.toLowerCase().lastIndexOf(firstName);
          reason = chunk.substring(0, nameIdx).trim();
          break;
        }
      }
    }
    
    // 3. Try to match any substring of student name
    if (!matchedStudent) {
      for (const student of searchPool) {
        const studentNameNorm = student.name.toLowerCase().trim();
        if (normalizedChunk.includes(studentNameNorm)) {
          matchedStudent = student;
          matchType = 'substring';
          const nameIdx = chunk.toLowerCase().indexOf(studentNameNorm);
          reason = (chunk.substring(0, nameIdx) + ' ' + chunk.substring(nameIdx + student.name.length)).trim();
          break;
        }
      }
    }
    
    // 4. Fallback: last word in chunk is treated as the name
    if (!matchedStudent && normalizedChunk.length > 0) {
      const words = chunk.split(/\s+/);
      const potentialName = words[words.length - 1];
      const potentialNameLower = potentialName.toLowerCase();
      const found = searchPool.find(s => s.name.toLowerCase().includes(potentialNameLower));
      if (found) {
        matchedStudent = found;
        matchType = 'fallback_match';
        reason = words.slice(0, words.length - 1).join(' ').trim();
      } else {
        matchedStudent = {
          id: null,
          name: potentialName,
          class: currentClass || 's2b'
        };
        matchType = 'unmatched';
        reason = words.slice(0, words.length - 1).join(' ').trim();
      }
    }
    
    if (reason) {
      reason = reason.replace(/^[^a-zA-Z0-9]+/, '').replace(/[^a-zA-Z0-9]+$/, '');
      reason = reason.replace(/^(note|and|for|with|due\s+to)\s+/i, '');
      reason = reason.trim();
    }
    
    if (matchedStudent) {
      results.push({
        student: { ...matchedStudent },
        amount: currentMatch.amount,
        type: currentMatch.type,
        reason: reason || 'Class activity',
        matchType
      });
    }
  }
  
  return results;
};

export default function App() {
  const [showSplash, setShowSplash] = useState(true);
  const [classToRemoveConfirm, setClassToRemoveConfirm] = useState(null);
  
  const [CLASSES, setCLASSES] = useState(() => {
    const stored = localStorage.getItem('caliph_classes');
    return stored ? JSON.parse(stored) : ['s2b', 'c2b', 'c2c', 's1b', 'c1b', 'c1c', 'jr8', 'jr9'];
  });
  const [hiddenClasses, setHiddenClasses] = useState(() => {
    const stored = localStorage.getItem('caliph_hidden_classes');
    return stored ? JSON.parse(stored) : [];
  });
  // --- JWT Auth State ---
  const [currentUser, setCurrentUser] = useState(() => {
    try { return JSON.parse(localStorage.getItem('caliph_user')); } catch { return null; }
  });
  const [authLoading, setAuthLoading] = useState(true);

  useEffect(() => {
    const token = localStorage.getItem('caliph_token');
    if (!token) { setAuthLoading(false); return; }
    fetch('/api/auth/verify', { headers: { Authorization: `Bearer ${token}` } })
      .then(r => r.ok ? r.json() : null)
      .then(data => {
        if (data?.user) setCurrentUser(data.user);
        else { localStorage.removeItem('caliph_token'); localStorage.removeItem('caliph_user'); setCurrentUser(null); }
      })
      .catch(() => {})
      .finally(() => setAuthLoading(false));
  }, []);

  const handleLogout = () => {
    localStorage.removeItem('caliph_token');
    localStorage.removeItem('caliph_user');
    setCurrentUser(null);
  };

  const getAuthHeader = () => ({ Authorization: `Bearer ${localStorage.getItem('caliph_token')}` });

  const hasPermission = (perm) => {
    if (!currentUser) return false;
    if (currentUser.role === 'super_admin' || currentUser.role === 'admin') return true;
    const perms = currentUser.permissions || [];
    return perms.includes('all') || perms.includes(perm);
  };

  useEffect(() => {
    localStorage.setItem('caliph_classes', JSON.stringify(CLASSES));
  }, [CLASSES]);
  useEffect(() => {
    localStorage.setItem('caliph_hidden_classes', JSON.stringify(hiddenClasses));
  }, [hiddenClasses]);

  const handleAddClass = () => {
    const newClass = prompt("Enter new class name (e.g. S3B):");
    if (newClass && newClass.trim() !== '') {
      const formattedName = newClass.trim().toLowerCase();
      if (!CLASSES.includes(formattedName)) {
        setCLASSES([...CLASSES, formattedName]);
      } else {
        alert("Class already exists!");
      }
    }
  };



  const handleEditClass = (oldName) => {
    const newName = prompt(`Enter new name for class ${oldName.toUpperCase()}:`, oldName.toUpperCase());
    if (newName && newName.trim() !== '') {
      const formattedNewName = newName.trim().toLowerCase();
      if (formattedNewName === oldName) return; // no change
      if (CLASSES.includes(formattedNewName)) {
        alert("A class with that name already exists!");
        return;
      }
      if (confirm(`Are you sure you want to rename ${oldName.toUpperCase()} to ${formattedNewName.toUpperCase()}? This will also update all students and mentors in this class.`)) {
        setCLASSES(CLASSES.map(c => c === oldName ? formattedNewName : c));
        
        if (hiddenClasses.includes(oldName)) {
          setHiddenClasses(hiddenClasses.map(c => c === oldName ? formattedNewName : c));
        }
        
        setStudents(students.map(s => {
          if (s.class === oldName) return { ...s, class: formattedNewName };
          return s;
        }));
        
        setMentors(mentors.map(m => {
          if (m.classAssigned === oldName) return { ...m, classAssigned: formattedNewName };
          return m;
        }));
      }
    }
  };
  
  const moveClassUp = (idx) => {
    if (idx === 0) return;
    const newClasses = [...CLASSES];
    const temp = newClasses[idx - 1];
    newClasses[idx - 1] = newClasses[idx];
    newClasses[idx] = temp;
    setCLASSES(newClasses);
  };

  const moveClassDown = (idx) => {
    if (idx === CLASSES.length - 1) return;
    const newClasses = [...CLASSES];
    const temp = newClasses[idx + 1];
    newClasses[idx + 1] = newClasses[idx];
    newClasses[idx] = temp;
    setCLASSES(newClasses);
  };

  const toggleClassVisibility = (cls) => {
    if (hiddenClasses.includes(cls)) {
      setHiddenClasses(hiddenClasses.filter(c => c !== cls));
    } else {
      setHiddenClasses([...hiddenClasses, cls]);
    }
  };
  
  useEffect(() => {
    const timer = setTimeout(() => {
      setShowSplash(false);
    }, 1500);
    return () => clearTimeout(timer);
  }, []);

  const [students, setStudents] = useState(INITIAL_STUDENTS);

  // Auto-synchronize CLASSES with student records so new classes from DB or imports automatically show up
  useEffect(() => {
    if (students && students.length > 0) {
      const studentClasses = Array.from(
        new Set(students.map(s => String(s.class || '').toLowerCase().trim()).filter(Boolean))
      );
      setCLASSES(prev => {
        const uniqueCombined = Array.from(new Set([...prev, ...studentClasses]));
        if (uniqueCombined.length !== prev.length || !uniqueCombined.every((val, idx) => val === prev[idx])) {
          return uniqueCombined;
        }
        return prev;
      });
    }
  }, [students]);
  
  const [mentors, setMentors] = useState(() => {
    const saved = localStorage.getItem('student_tracker_mentors');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return INITIAL_MENTORS;
  });
  
  useEffect(() => {
    localStorage.setItem('student_tracker_mentors', JSON.stringify(mentors));
  }, [mentors]);

  const [showAddMentorModal, setShowAddMentorModal] = useState(false);
  const [newMentor, setNewMentor] = useState({ name: '', role: '', classAssigned: 's2b', email: '', rating: 5.0, bio: '', roomNumber: '' });
  
  const [showDeleteMentorModal, setShowDeleteMentorModal] = useState(false);
  const [selectedMentorForDelete, setSelectedMentorForDelete] = useState(null);

  const handleAddMentor = (e) => {
    e.preventDefault();
    if (!newMentor.name.trim() || !newMentor.roomNumber.trim()) return;
    const newM = {
      ...newMentor,
      id: 'm' + Date.now()
    };
    setMentors([...mentors, newM]);
    setNewMentor({ name: '', role: '', classAssigned: 's2b', email: '', rating: 5.0, bio: '', roomNumber: '' });
    setShowAddMentorModal(false);
  };

  const handleDeleteMentor = (e) => {
    e.preventDefault();
    if (!selectedMentorForDelete) return;
    
    // Find the mentor name to clean up mentorStudents if needed
    const mentorToDelete = mentors.find(m => m.id === selectedMentorForDelete);
    if (mentorToDelete) {
      setMentorStudents(prev => {
        const next = { ...prev };
        delete next[mentorToDelete.name];
        return next;
      });
    }

    setMentors(mentors.filter(m => m.id !== selectedMentorForDelete));
    setSelectedMentorForDelete(null);
    setShowDeleteMentorModal(false);
  };
  const [performanceView, setPerformanceView] = useState(null); // 'neat', 'room', 'ineligible', null
  const [selectedHostel, setSelectedHostel] = useState(null); // 'MAIN BLOCK BOYS', 'NEW BLOCK BOYS', 'MAIN BLOCK GIRLS', 'NEW BLOCK GIRLS'
  const [selectedRoom, setSelectedRoom] = useState(null); // 'Room 1' .. 'Room 25'
  const HOSTEL_BLOCKS = [
    { id: 1, name: 'MAIN BLOCK BOYS', icon: '👦', desc: 'Main Boys Hostel' },
    { id: 2, name: 'NEW BLOCK BOYS', icon: '🏢', desc: 'New Boys Hostel' },
    { id: 3, name: 'MAIN BLOCK GIRLS', icon: '👧', desc: 'Main Girls Hostel' },
    { id: 4, name: 'NEW BLOCK GIRLS', icon: '🏠', desc: 'New Girls Hostel' }
  ];
  const [performanceSelectedClass, setPerformanceSelectedClass] = useState(null);
  const [performanceSelectedStudents, setPerformanceSelectedStudents] = useState([]);
  const [showPerformanceSubmitModal, setShowPerformanceSubmitModal] = useState(false);
  const [performanceSubmitData, setPerformanceSubmitData] = useState({ count: 1, reason: '', type: 'tally' });
  const [showPerformanceReasonDropdown, setShowPerformanceReasonDropdown] = useState(false);
  const [showWhatsappModal, setShowWhatsappModal] = useState(false);
  const [whatsappMessage, setWhatsappMessage] = useState('');
  
  // Room Tally & Fine workflow states
  const [showRoomTallyModal, setShowRoomTallyModal] = useState(false);
  const [roomTallyCount, setRoomTallyCount] = useState(1);
  const [roomTallyReason, setRoomTallyReason] = useState('');
  const [roomTallySummary, setRoomTallySummary] = useState(null);

  const [showRoomFineModal, setShowRoomFineModal] = useState(false);
  const [roomFineAmount, setRoomFineAmount] = useState('');
  const [roomFineReason, setRoomFineReason] = useState('');
  const [roomFineSummary, setRoomFineSummary] = useState(null);
  
  const [showAddIneligibleModal, setShowAddIneligibleModal] = useState(false);
  const [ineligibleSelectedStudents, setIneligibleSelectedStudents] = useState([]);
  const [ineligibleReasonInput, setIneligibleReasonInput] = useState('');
  
  const [spotClass, setSpotClass] = useState('');
  const [spotNameSearch, setSpotNameSearch] = useState('');
  const [spotReason, setSpotReason] = useState('');
  const [spotAmount, setSpotAmount] = useState('');
  const [showSpotDropdown, setShowSpotDropdown] = useState(false);
  const [showSpotReasonDropdown, setShowSpotReasonDropdown] = useState(false);

  const [programClass, setProgramClass] = useState('');
  const [programNameSearch, setProgramNameSearch] = useState('');
  const [programReason, setProgramReason] = useState('');
  const [programAmount, setProgramAmount] = useState('');
  const [showProgramDropdown, setShowProgramDropdown] = useState(false);

  const [sheetClass, setSheetClass] = useState('');
  const [sheetNameSearch, setSheetNameSearch] = useState('');
  const [sheetReason, setSheetReason] = useState('');
  const [showSheetDropdown, setShowSheetDropdown] = useState(false);

  // Morning Bliss State
  const [morningBlissClass, setMorningBlissClass] = useState('');
  const [morningBlissNameSearch, setMorningBlissNameSearch] = useState('');
  const [morningBlissTopic, setMorningBlissTopic] = useState('');
  const [morningBlissMark, setMorningBlissMark] = useState('');
  const [morningBlissEv, setMorningBlissEv] = useState('');
  const [showMorningBlissDropdown, setShowMorningBlissDropdown] = useState(false);
  const [morningBlissDuration, setMorningBlissDuration] = useState(0); // seconds
  const [isStopwatchRunning, setIsStopwatchRunning] = useState(false);
  const [mbFromDate, setMbFromDate] = useState(() => new Date().toISOString().split('T')[0]);
  const handleDownloadScreenshotCard = async (elementId, fileName = 'Report_Card.png') => {
    try {
      const cardElement = document.getElementById(elementId);
      if (!cardElement) {
        alert('Card element not found for screenshot!');
        return;
      }
      const canvas = await html2canvas(cardElement, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#FFFFFF',
        logging: false
      });
      const imageUri = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      link.href = imageUri;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error('Error generating screenshot card:', err);
      alert('Failed to generate screenshot image. Please try again.');
    }
  };

  useEffect(() => {
    let interval;
    if (isStopwatchRunning) {
      interval = setInterval(() => {
        setMorningBlissDuration(prev => prev + 1);
      }, 1000);
    } else if (!isStopwatchRunning && morningBlissDuration !== 0) {
      clearInterval(interval);
    }
    return () => clearInterval(interval);
  }, [isStopwatchRunning, morningBlissDuration]);

  const formatMBTime = (totalSeconds) => {
    const m = Math.floor(totalSeconds / 60).toString().padStart(2, '0');
    const s = (totalSeconds % 60).toString().padStart(2, '0');
    return `${m}:${s}`;
  };

  const handleMorningBlissExportExcelRange = () => {
    const tableData = [];
    const star3 = [];
    const star2 = [];
    const star1 = [];

    students.forEach(s => {
      if ((s.morningBlissMark != null && s.morningBlissMark !== '') || s.morningBlissScript != null) {
        if (s.summaryDate && s.summaryDate >= mbFromDate && s.summaryDate <= mbToDate) {
          const total = (Number(s.morningBlissMark) || 0) + (Number(s.morningBlissScript) || 0);
          tableData.push([
            s.class.toUpperCase(),
            s.name,
            s.morningBlissTopic || '-',
            total,
            s.morningBlissEv || '-',
            s.summaryDate
          ]);

          if (total === 10) star3.push(s.name);
          else if (total >= 9.5) star2.push(s.name);
          else if (total >= 9) star1.push(s.name);
        }
      }
    });

    if (tableData.length === 0) {
      alert(`No Morning Bliss records found with Summary date between ${mbFromDate} and ${mbToDate}`);
      return;
    }

    const excelData = [
      ['CLASS', 'NAME', 'TOPIC', 'SCORE', 'EVALUATED', 'SUMMARY DATE'],
      ...tableData,
      [],
      ['3 Stars 🌟🌟🌟'],
      ...star3.map(n => [n]),
      [],
      ['2 Stars 🌟🌟'],
      ...star2.map(n => [n]),
      [],
      ['1 Star 🌟'],
      ...star1.map(n => [n])
    ];

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Morning Bliss");
    XLSX.writeFile(wb, `Morning_Bliss_Summary_${mbFromDate}_to_${mbToDate}.xlsx`);
  };

  const handleMorningBlissSubmit = async (e) => {
    if (e) e.preventDefault();
    if (!morningBlissClass || !morningBlissNameSearch || !morningBlissTopic || !morningBlissMark || !morningBlissEv) return;

    const match = students.find(s => s.class === morningBlissClass && s.name.toLowerCase() === morningBlissNameSearch.toLowerCase());
    const studentToUpdate = match || students.find(s => s.class === morningBlissClass && s.name.toLowerCase().startsWith(morningBlissNameSearch.toLowerCase()));

    if (!studentToUpdate) {
        alert("Student not found!");
        return;
    }
    
    const mark = Number(morningBlissMark);
    if (mark > 8) {
      alert("Maximum mark is 8!");
      return;
    }

    const todayStr = new Date().toISOString().split('T')[0];
    let summaryId = studentToUpdate.summaryId;
    let summaryDate = todayStr;

    try {
      const res = await fetch('/api/morning-bliss/summary', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date: todayStr })
      });
      if (res.ok) {
        const sumData = await res.json();
        summaryId = sumData.id;
        summaryDate = sumData.date;
      }
    } catch (err) {
      console.error("Failed to fetch/create summary", err);
    }
    
    // If student's summaryDate is changing to today, reset script mark & MB stars for today's fresh evaluation
    const isNewDay = studentToUpdate.summaryDate !== todayStr;
    const updated = {
        ...studentToUpdate,
        morningBlissMark: mark,
        morningBlissTopic: morningBlissTopic,
        morningBlissEv: morningBlissEv,
        morningBlissDuration: formatMBTime(morningBlissDuration),
        morningBlissScript: isNewDay ? null : studentToUpdate.morningBlissScript,
        morningBlissStar: isNewDay ? 0 : studentToUpdate.morningBlissStar,
        summaryId: summaryId,
        summaryDate: summaryDate
    };
    
    setStudents(prev => prev.map(s => s.id === studentToUpdate.id ? updated : s));
    
    fetch('/api/students/bulk-upsert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ students: [updated] })
    }).catch(err => console.error("Error bulk upserting:", err));

    const now = new Date();
    const formattedDate = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    
    const summary = `*MORNING BLISS REPORT*
Date: ${formattedDate}
Class: ${morningBlissClass.toUpperCase()}
Name: ${studentToUpdate.name}
Topic: ${morningBlissTopic}
Duration: ${formatMBTime(morningBlissDuration)}
Score: ${mark}
EV: ${morningBlissEv}`;

    setWhatsappMessage(summary);
    setShowWhatsappModal(true);
    
    setMorningBlissNameSearch('');
    setMorningBlissMark('');
    setMorningBlissDuration(0);
    setIsStopwatchRunning(false);
  };

  const handleMorningBlissDone = async () => {
    // Generate a professional PNG star from an SVG to embed in the PDF
    const starImgData = await new Promise((resolve) => {
      const starSvg = '<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 24 24" fill="#EAB308" stroke="#CA8A04" stroke-width="1"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>';
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = 64;
        canvas.height = 64;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        resolve(canvas.toDataURL('image/png'));
      };
      img.src = 'data:image/svg+xml;base64,' + btoa(starSvg);
    });

    const doc = new jsPDF();
    const now = new Date();
    const formattedDate = now.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    const todayDateStr = now.toISOString().split('T')[0];
    
    // --- Data Prep (Strictly Today's Data based on Summary Date) ---
    const tableData = [];
    const star3 = [];
    const star2 = [];
    const star1 = [];
    
    CLASSES.forEach(c => {
      const classStudents = students.filter(s => s.class === c && s.morningBlissMark != null && s.morningBlissMark !== '' && (s.summaryDate === todayDateStr || !s.summaryDate));
      classStudents.forEach(s => {
        const total = (Number(s.morningBlissMark) || 0) + (Number(s.morningBlissScript) || 0);
        tableData.push([
          s.class.toUpperCase(),
          s.name,
          s.morningBlissTopic || '-',
          total,
          s.morningBlissEv || '-'
        ]);
        
        let newStars = 0;
        if (total === 10) {
          star3.push(s.name);
          newStars = 3;
        } else if (total >= 9.5) {
          star2.push(s.name);
          newStars = 2;
        } else if (total >= 9) {
          star1.push(s.name);
          newStars = 1;
        }

        if (newStars > 0) {
          logHistory(s.id, 'Morning Bliss', newStars, s.morningBlissTopic || 'Morning Bliss Evaluation');
        }
      });
    });

    // --- Generate PDF ---
    doc.setFontSize(14);
    doc.text(`Morning Bliss Results - ${formattedDate}`, 105, 15, null, null, "center");

    autoTable(doc, {
      startY: 25,
      head: [['CLASS', 'NAME', 'TOPIC', 'SCORE', 'EVALUDED']],
      body: tableData,
      theme: 'grid',
      headStyles: { fillColor: [26, 54, 93] },
      styles: { fontSize: 10, cellPadding: 3 }
    });

    let currentY = doc.lastAutoTable.finalY + 15;
    
    doc.setFontSize(12);
    doc.setFont("helvetica", "bold");
    
    if (star3.length > 0) {
      doc.text("3 Stars", 14, currentY);
      doc.addImage(starImgData, 'PNG', 32, currentY - 4, 5, 5);
      doc.addImage(starImgData, 'PNG', 37.5, currentY - 4, 5, 5);
      doc.addImage(starImgData, 'PNG', 43, currentY - 4, 5, 5);
      currentY += 8;
      doc.setFont("helvetica", "normal");
      star3.forEach(name => {
        doc.text(`- ${name}`, 20, currentY);
        currentY += 6;
      });
      currentY += 4;
      doc.setFont("helvetica", "bold");
    }
    
    if (star2.length > 0) {
      doc.text("2 Stars", 14, currentY);
      doc.addImage(starImgData, 'PNG', 32, currentY - 4, 5, 5);
      doc.addImage(starImgData, 'PNG', 37.5, currentY - 4, 5, 5);
      currentY += 8;
      doc.setFont("helvetica", "normal");
      star2.forEach(name => {
        doc.text(`- ${name}`, 20, currentY);
        currentY += 6;
      });
      currentY += 4;
      doc.setFont("helvetica", "bold");
    }
    
    if (star1.length > 0) {
      doc.text("1 Star", 14, currentY);
      doc.addImage(starImgData, 'PNG', 30, currentY - 4, 5, 5);
      currentY += 8;
      doc.setFont("helvetica", "normal");
      star1.forEach(name => {
        doc.text(`- ${name}`, 20, currentY);
        currentY += 6;
      });
    }

    doc.save(`Morning_Bliss_Summary_${formattedDate}.pdf`);

    // --- Generate Excel ---
    const excelData = [
      ['CLASS', 'NAME', 'TOPIC', 'SCORE', 'EVALUDED'],
      ...tableData,
      [],
      ['3 Stars 🌟🌟🌟'],
      ...star3.map(n => [n]),
      [],
      ['2 Stars 🌟🌟'],
      ...star2.map(n => [n]),
      [],
      ['1 Star 🌟'],
      ...star1.map(n => [n])
    ];

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Morning Bliss");
    XLSX.writeFile(wb, `Morning_Bliss_Summary_${formattedDate}.xlsx`);
  };


  const handleSheetSubmit = (e) => {
    e.preventDefault();
    if (!sheetClass || !sheetNameSearch || !sheetReason || !performanceView) return;

    let typeLabel = '';
    if (performanceView === 'sheets_black') typeLabel = 'Black Sheet';
    if (performanceView === 'sheets_yellow') typeLabel = 'Yellow Sheet';
    if (performanceView === 'sheets_apology') typeLabel = 'Apology';

    const match = students.find(s => s.class === sheetClass && s.name.toLowerCase() === sheetNameSearch.toLowerCase());
    const studentToUpdate = match || students.find(s => s.class === sheetClass && s.name.toLowerCase().startsWith(sheetNameSearch.toLowerCase()));

    if (!studentToUpdate) {
        alert("Student not found!");
        return;
    }

    if (performanceView === 'sheets_black') {
      updateStudentField(studentToUpdate.id, 'sheetTally', (studentToUpdate.sheetTally || 0) - 8);
      updateStudentField(studentToUpdate.id, 'ineligible', true);
      updateStudentField(studentToUpdate.id, 'ineligibleReason', sheetReason ? `Black Sheet - ${sheetReason}` : 'Black Sheet');
      logHistory(studentToUpdate.id, 'Black Sheet', -8, sheetReason);
    } else if (performanceView === 'sheets_yellow') {
      updateStudentField(studentToUpdate.id, 'sheetTally', (studentToUpdate.sheetTally || 0) - 4);
      logHistory(studentToUpdate.id, 'Yellow Sheet', -4, sheetReason);
    } else if (performanceView === 'sheets_apology') {
      updateStudentField(studentToUpdate.id, 'sheetTally', (studentToUpdate.sheetTally || 0) - 1.5);
      logHistory(studentToUpdate.id, 'Apology Sheet', -1.5, sheetReason);
    }

    setSheetNameSearch('');
    setSheetReason('');

    const now = new Date();
    const formattedDate = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    const formattedTime = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    
    const sheetSummary = `📄 *SHEET REPORT*
Type: *${typeLabel}*
Date: ${formattedDate}
Time: ${formattedTime}
Class: CLASS ${sheetClass.toUpperCase()}
Name: ${studentToUpdate.name}
Reason: ${sheetReason}`;

    setSummaryText(sheetSummary);
    setShowSummaryModal(true);
  };

  const handleProgramStarSubmit = (e) => {
    e.preventDefault();
    if (!programClass || !programNameSearch || !programAmount) return;

    const match = students.find(s => s.class === programClass && s.name.toLowerCase() === programNameSearch.toLowerCase());
    const studentToUpdate = match || students.find(s => s.class === programClass && s.name.toLowerCase().startsWith(programNameSearch.toLowerCase()));

    if (!studentToUpdate) {
        alert("Student not found!");
        return;
    }
    
    const amount = Number(programAmount);
    const updated = {
        ...studentToUpdate,
        star: (studentToUpdate.star || 0) + amount,
        programmeReason: programReason || studentToUpdate.programmeReason
    };
    
    setStudents(students.map(s => s.id === studentToUpdate.id ? updated : s));
    
    fetch('/api/students/bulk-upsert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ students: [updated] })
    }).catch(err => console.error("Error bulk upserting:", err));

    setProgramNameSearch('');
    setProgramReason('');
    setProgramAmount('');
    
    const now = new Date();
    const formattedDate = now.toLocaleDateString();
    const formattedTime = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const programStarSummary = `Date: ${formattedDate}
Time: ${formattedTime}
Class: ${programClass.toUpperCase()}
Name: ${studentToUpdate.name}
Programme: ${programReason || 'N/A'}
Star Count: ${amount}`;

    setSummaryText(programStarSummary);
    setShowSummaryModal(true);
  };

  const handleSpotFineSubmit = (e) => {
    e.preventDefault();
    if (!spotClass || !spotNameSearch || !spotAmount) return;

    const match = students.find(s => s.class === spotClass && s.name.toLowerCase() === spotNameSearch.toLowerCase());
    const studentToUpdate = match || students.find(s => s.class === spotClass && s.name.toLowerCase().startsWith(spotNameSearch.toLowerCase()));

    if (!studentToUpdate) {
        alert("Student not found!");
        return;
    }
    
    const amount = Number(spotAmount);
    const updated = {
        ...studentToUpdate,
        fine: (studentToUpdate.fine || 0) + amount,
        fineCount: (studentToUpdate.fineCount || 0) + 1,
        fineReason: spotReason || studentToUpdate.fineReason
    };
    
    logHistory(studentToUpdate.id, 'Spot Fine', amount, spotReason);
    
    setStudents(students.map(s => s.id === studentToUpdate.id ? updated : s));
    
    fetch('/api/students/bulk-upsert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ students: [updated] })
    }).catch(err => console.error("Error bulk upserting:", err));

    setSpotNameSearch('');
    setSpotReason('');
    setSpotAmount('');
    
    const now = new Date();
    const formattedDate = now.toLocaleDateString();
    const formattedTime = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const spotFineSummary = `Date: ${formattedDate}
Time: ${formattedTime}
Class: ${spotClass.toUpperCase()}
Name: ${studentToUpdate.name}
Reason: ${spotReason}
Fine Amount: ${amount}`;

    setSummaryText(spotFineSummary);
    setShowSummaryModal(true);
  };
  
  const handlePerformanceSubmit = (e) => {
    e.preventDefault();
    const amount = Number(performanceSubmitData.count);
    const reason = performanceSubmitData.reason;

    let studentsToUpsert = [];
    let messageLines = [`*${performanceView.toUpperCase()} TALLY REPORT*`, `Reason: ${reason}`, ''];
    
    const updatedStudents = students.map(s => {
      if (performanceSelectedStudents.includes(s.id)) {
        const isNeat = performanceView === 'neat';
        const updated = {
          ...s,
          tally: isNeat ? s.tally : (s.tally || 0) + amount,
          neatAndOrderTally: isNeat ? ((s.neatAndOrderTally || 0) + amount) : (s.neatAndOrderTally || 0),
          neatAndOrderIncidents: isNeat ? ((s.neatAndOrderIncidents || 0) + 1) : (s.neatAndOrderIncidents || 0),
          tallyReason: isNeat ? s.tallyReason : (reason || s.tallyReason),
          neatAndOrderReason: isNeat ? (reason || s.neatAndOrderReason) : s.neatAndOrderReason
        };
        logHistory(s.id, isNeat ? 'N&O' : 'tally', isNeat ? 1 : amount, reason);
        studentsToUpsert.push(updated);
        messageLines.push(`- ${s.name} (Class ${s.class.toUpperCase()}): ${amount} tally`);
        return updated;
      }
      return s;
    });

    setStudents(updatedStudents);
    
    fetch('/api/students/bulk-upsert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ students: studentsToUpsert })
    }).catch(err => console.error("Error bulk upserting:", err));

    setWhatsappMessage(messageLines.join('\n'));
    setShowPerformanceSubmitModal(false);
    setShowWhatsappModal(true);
  };
  
  const handleAddIneligibleSubmit = (e) => {
    e.preventDefault();
    if (ineligibleSelectedStudents.length === 0) return;
    
    let studentsToUpsert = [];
    const updatedStudents = students.map(s => {
      if (ineligibleSelectedStudents.includes(s.id)) {
        const updated = {
          ...s,
          ineligible: true,
          ineligibleReason: ineligibleReasonInput || s.ineligibleReason
        };
        studentsToUpsert.push(updated);
        return updated;
      }
      return s;
    });

    setStudents(updatedStudents);
    fetch('/api/students/bulk-upsert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ students: studentsToUpsert })
    }).catch(err => console.error("Error bulk upserting ineligible:", err));

    setShowAddIneligibleModal(false);
    setIneligibleSelectedStudents([]);
    setIneligibleReasonInput('');
  };
  
  const handleRemoveIneligible = (studentId) => {
    let studentToUpdate = null;
    const updatedStudents = students.map(s => {
      if (s.id === studentId) {
        studentToUpdate = { ...s, ineligible: false, ineligibleReason: '' };
        return studentToUpdate;
      }
      return s;
    });
    if (studentToUpdate) {
      setStudents(updatedStudents);
      fetch('/api/students/bulk-upsert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ students: [studentToUpdate] })
      }).catch(err => console.error("Error bulk upserting ineligible:", err));
    }
  };
  
  const [dbLoading, setDbLoading] = useState(true);
  const [dbError, setDbError] = useState(null);
  
  // Tab control: 'scoring' or 'admin'
  const [activeTab, setActiveTab] = useState('scoring');

  const [selectedClass, setSelectedClass] = useState(null);
  const [selectedStudentIds, setSelectedStudentIds] = useState([]);
  const [isScoring, setIsScoring] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [showAddModal, setShowAddModal] = useState(false);
  const [newStudentName, setNewStudentName] = useState('');
  const [addStudentMethod, setAddStudentMethod] = useState('single'); // 'single' or 'excel'
  
  // Admin sheet is gated by role check from JWT
  const isSuperAdmin = currentUser?.role === 'super_admin' || currentUser?.username === 'admin';
  const isAdminAuthenticated = isSuperAdmin || currentUser?.role === 'admin';
  const [passwordError, setPasswordError] = useState('');
  // User management state
  const [adminSubTab, setAdminSubTab] = useState('sheet'); // 'sheet' | 'users'
  const [usersList, setUsersList] = useState([]);
  const [usersLoading, setUsersLoading] = useState(false);
  const [showUserModal, setShowUserModal] = useState(false);
  const [editingUser, setEditingUser] = useState(null);
  const [userForm, setUserForm] = useState({ username: '', password: '', role: 'user', permissions: [] });
  const [userFormError, setUserFormError] = useState('');

  const ALL_PERMISSIONS = [
    { key: 'scoring', label: 'Scoring' },
    { key: 'attendance', label: 'Attendance' },
    { key: 'morning_bliss', label: 'Morning Bliss' },
    { key: 'performance', label: 'Performance' },
    { key: 'mentor', label: 'Mentor' },
    { key: 'admin_sheet', label: 'Admin Sheet' },
    { key: 'view_report', label: 'View Reports' },
    { key: 'spot_fine', label: 'Spot Fine' },
    { key: 'ineligible', label: 'Ineligible' },
  ];

  const fetchUsers = async () => {
    setUsersLoading(true);
    try {
      const res = await fetch('/api/users', { headers: getAuthHeader() });
      if (res.ok) setUsersList(await res.json());
    } catch {} finally { setUsersLoading(false); }
  };

  useEffect(() => {
    if (adminSubTab === 'users' && isSuperAdmin) fetchUsers();
  }, [adminSubTab]);

  const handleSaveUser = async (e) => {
    e.preventDefault();
    setUserFormError('');
    if (!userForm.username.trim() || (!editingUser && !userForm.password.trim())) {
      setUserFormError('Username and password are required.'); return;
    }
    const method = editingUser ? 'PUT' : 'POST';
    const url = editingUser ? `/api/users/${editingUser.id}` : '/api/users';
    try {
      const res = await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json', ...getAuthHeader() },
        body: JSON.stringify(userForm),
      });
      const data = await res.json();
      if (!res.ok) { setUserFormError(data.error || 'Failed to save user.'); return; }
      setShowUserModal(false);
      setEditingUser(null);
      setUserForm({ username: '', password: '', role: 'user', permissions: [] });
      fetchUsers();
    } catch { setUserFormError('Connection error.'); }
  };

  const handleDeleteUser = async (id) => {
    if (!window.confirm('Delete this user?')) return;
    await fetch(`/api/users/${id}`, { method: 'DELETE', headers: getAuthHeader() });
    fetchUsers();
  };
  const [showDownloadModal, setShowDownloadModal] = useState(false);
  const [downloadSelectedClasses, setDownloadSelectedClasses] = useState([]);
  const [showClearModal, setShowClearModal] = useState(false);
  const [clearSelectedClasses, setClearSelectedClasses] = useState([]);
  const [showAdminSettingsModal, setShowAdminSettingsModal] = useState(false);
  
  const visibleClasses = CLASSES.filter(c => !hiddenClasses.includes(c));

  // Admin selected class dropdown
  const [adminClass, setAdminClass] = useState(() => visibleClasses[0] || 's2b');
  const [adminSelectedStudents, setAdminSelectedStudents] = useState([]);

  const handleAdminDeleteSelectedStudents = () => {
    if (adminSelectedStudents.length === 0) {
      alert("Please select at least one student to delete.");
      return;
    }
    if (window.confirm(`Are you sure you want to delete ${adminSelectedStudents.length} selected student(s)?`)) {
      adminSelectedStudents.forEach(studentId => {
        fetch(`/api/students/${studentId}`, { method: 'DELETE' }).catch(err => console.error("Error deleting:", err));
      });
      setStudents(prev => prev.filter(s => !adminSelectedStudents.includes(s.id)));
      setAdminSelectedStudents([]);
    }
  };



  // Mentee Rooms & Teacher Mentorship Dashboard states
  const [menteeRooms, setMenteeRooms] = useState(() => {
    const saved = localStorage.getItem('student_tracker_mentee_rooms');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return [
      {
        id: 'room-1',
        name: 'STEM Guidance & Science Club',
        mentorName: 'Prof. Rashid Ali',
        classTarget: 's2b',
        studentIds: ['s2b-1', 's2b-2', 's2b-4'],
        description: 'Focus on Science Olympiad prep & advanced lab work.',
        createdAt: '2026-07-21'
      },
      {
        id: 'room-2',
        name: 'Cambridge Leadership Mentees',
        mentorName: 'Dr. Ananya Sharma',
        classTarget: 'c2b',
        studentIds: ['c2b-1', 'c2b-4'],
        description: 'Public speaking & global curriculum mentorship.',
        createdAt: '2026-07-21'
      }
    ];
  });

  const [showCreateRoomModal, setShowCreateRoomModal] = useState(false);
  const [newRoomName, setNewRoomName] = useState('');
  const [newRoomDesc, setNewRoomDesc] = useState('');
  const [newRoomClass, setNewRoomClass] = useState('s2b');
  const [newRoomMentor, setNewRoomMentor] = useState('Prof. Rashid Ali');

  const [showAddStudentToRoomModal, setShowAddStudentToRoomModal] = useState(false);
  const [targetRoomId, setTargetRoomId] = useState(null);
  const [selectedStudentForRoom, setSelectedStudentForRoom] = useState('');

  // Mentor switch selector & sub-tabs
  const [selectedMentorIndex, setSelectedMentorIndex] = useState(null);
  const [mentorSubTab, setMentorSubTab] = useState('diary'); // 'diary' | 'attendance'

  // Night Attendance state
  const [attendanceRecords, setAttendanceRecords] = useState({});
  const [showAttendanceSummary, setShowAttendanceSummary] = useState(false);

  const handleAttendanceChange = (studentId, field, value) => {
    setAttendanceRecords(prev => ({
      ...prev,
      [studentId]: { ...prev[studentId], [field]: value }
    }));
  };

  // Mentor Diary state
  const [diaryRecords, setDiaryRecords] = useState({});
  const [showDiarySummary, setShowDiarySummary] = useState(false);

  const handleDiaryChange = (studentId, field, value) => {
    setDiaryRecords(prev => ({
      ...prev,
      [studentId]: { ...prev[studentId], [field]: value }
    }));
  };

  // Mentor-Student assignment state
  const [mentorStudents, setMentorStudents] = useState(() => {
    const saved = localStorage.getItem('student_tracker_mentor_students');
    return saved ? JSON.parse(saved) : {};
  });
  const [showMentorAddStudentModal, setShowMentorAddStudentModal] = useState(false);
  const [selectedStudentsForMentor, setSelectedStudentsForMentor] = useState([]);
  const [mentorStudentSearch, setMentorStudentSearch] = useState('');
  const [showMentorRemoveStudentModal, setShowMentorRemoveStudentModal] = useState(false);
  const [selectedStudentForMentorRemove, setSelectedStudentForMentorRemove] = useState('');

  const handleRemoveStudentFromMentor = (e) => {
    e.preventDefault();
    if (!selectedStudentForMentorRemove || selectedMentorIndex === null) return;
    const mentor = mentors[selectedMentorIndex];
    setMentorStudents(prev => {
      const current = prev[mentor.name] || [];
      return { ...prev, [mentor.name]: current.filter(id => id !== selectedStudentForMentorRemove) };
    });
    setSelectedStudentForMentorRemove('');
    setShowMentorRemoveStudentModal(false);
  };

  // Sync menteeRooms and mentorStudents to localStorage
  useEffect(() => {
    localStorage.setItem('student_tracker_mentee_rooms', JSON.stringify(menteeRooms));
  }, [menteeRooms]);

  useEffect(() => {
    localStorage.setItem('student_tracker_mentor_students', JSON.stringify(mentorStudents));
  }, [mentorStudents]);

  const handleCreateRoom = (e) => {
    e.preventDefault();
    if (!newRoomName.trim()) return;
    const newRoom = {
      id: `room-${Date.now()}`,
      name: newRoomName.trim(),
      mentorName: newRoomMentor,
      classTarget: newRoomClass,
      studentIds: [],
      description: newRoomDesc.trim() || 'General Mentorship Room',
      createdAt: new Date().toISOString().split('T')[0]
    };
    setMenteeRooms(prev => [...prev, newRoom]);
    setNewRoomName('');
    setNewRoomDesc('');
    setShowCreateRoomModal(false);
  };

  const handleAddStudentToRoom = (e) => {
    e.preventDefault();
    if (!targetRoomId || !selectedStudentForRoom) return;
    setMenteeRooms(prev => prev.map(room => {
      if (room.id === targetRoomId) {
        if (room.studentIds.includes(selectedStudentForRoom)) return room;
        return { ...room, studentIds: [...room.studentIds, selectedStudentForRoom] };
      }
      return room;
    }));
    setSelectedStudentForRoom('');
    setShowAddStudentToRoomModal(false);
  };

  const handleAddStudentToMentor = (e) => {
    e.preventDefault();
    if (selectedStudentsForMentor.length === 0 || selectedMentorIndex === null) return;
    const mentor = mentors[selectedMentorIndex];
    setMentorStudents(prev => {
      const current = prev[mentor.name] || [];
      const newStudents = selectedStudentsForMentor.filter(id => !current.includes(id));
      if (newStudents.length === 0) return prev;
      return { ...prev, [mentor.name]: [...current, ...newStudents] };
    });
    setSelectedStudentsForMentor([]);
    setMentorStudentSearch('');
    setShowMentorAddStudentModal(false);
  };

  const handleRemoveStudentFromRoom = (roomId, studentId) => {
    setMenteeRooms(prev => prev.map(room => {
      if (room.id === roomId) {
        return { ...room, studentIds: room.studentIds.filter(id => id !== studentId) };
      }
      return room;
    }));
  };

  // Classroom sub-tabs: 'scoring' (student list), 'attendance' (Night Attendance), 'diary' (Daily Diary)
  const [roomTab, setRoomTab] = useState('scoring');

  // Night Attendance state map: { [studentId]: 'present' | 'absent' | 'leave' }
  const [attendanceMap, setAttendanceMap] = useState({});

  // Daily Diary entries state
  const [diaries, setDiaries] = useState(() => {
    const saved = localStorage.getItem('student_tracker_diaries');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return [
      {
        id: 'diary-1',
        class: 's2b',
        date: '2026-07-21',
        title: 'Science & Math Home Task',
        subject: 'Science / Math',
        content: 'Complete Chapter 4 exercise questions in notebook. Prepare for tomorrow\'s quiz on Newton\'s Laws of Motion.',
        teacherName: 'Prof. Rashid Ali'
      }
    ];
  });

  const [newDiarySubject, setNewDiarySubject] = useState('Science / Math');
  const [newDiaryContent, setNewDiaryContent] = useState('');
  const [diaryStartDate, setDiaryStartDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 7);
    return d.toISOString().split('T')[0];
  });
  const [diaryEndDate, setDiaryEndDate] = useState(() => new Date().toISOString().split('T')[0]);

  // WhatsApp Report Modal states
  const [showWhatsAppModal, setShowWhatsAppModal] = useState(false);
  const [whatsAppReportText, setWhatsAppReportText] = useState('');

  // Save diaries to localStorage
  useEffect(() => {
    localStorage.setItem('student_tracker_diaries', JSON.stringify(diaries));
  }, [diaries]);

  // Handle mobile hardware back button using hash history
  const isHome = activeTab === 'scoring' && !performanceView && !selectedClass && !showAdminSettingsModal && selectedMentorIndex === null;

  useEffect(() => {
    if (!isHome && window.location.hash !== '#menu') {
      window.history.pushState(null, '', '#menu');
    } else if (isHome && window.location.hash === '#menu') {
      window.history.back();
    }
  }, [isHome]);

  useEffect(() => {
    const handleHashChange = () => {
      if (window.location.hash !== '#menu' && !isHome) {
        if (showAdminSettingsModal) {
          setShowAdminSettingsModal(false);
        } else if (selectedMentorIndex !== null) {
          setSelectedMentorIndex(null);
        } else if (performanceView) {
          if (performanceView.startsWith('sheets_')) {
            setPerformanceView('sheets');
            window.history.pushState(null, '', '#menu');
          } else {
            setPerformanceView(null);
          }
        } else if (selectedClass) {
          setSelectedClass(null);
        } else if (activeTab !== 'scoring') {
          setActiveTab('scoring');
        }
      }
    };
    window.addEventListener('hashchange', handleHashChange);
    return () => window.removeEventListener('hashchange', handleHashChange);
  }, [performanceView, activeTab, isHome, showAdminSettingsModal, selectedMentorIndex, selectedClass]);

  // Automatically mark every student in the opened class as 'present' when class changes
  useEffect(() => {
    if (selectedClass && students.length > 0) {
      const classStus = students.filter(s => s.class === selectedClass);
      const initialMap = {};
      classStus.forEach(s => {
        initialMap[s.id] = 'present'; // Automatically Present!
      });
      setAttendanceMap(initialMap);
    }
  }, [selectedClass, students]);

  // Generate Night Attendance WhatsApp Report
  const generateWhatsAppAttendanceReport = () => {
    const classStus = students.filter(s => s.class === selectedClass);
    const presentStudents = classStus.filter(s => (attendanceMap[s.id] || 'present') === 'present');
    const absentStudents = classStus.filter(s => (attendanceMap[s.id] || 'present') === 'absent');
    const leaveStudents = classStus.filter(s => (attendanceMap[s.id] || 'present') === 'leave');

    const todayStr = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });

    let text = `🌙 *CALIPH SCHOOL - NIGHT ATTENDANCE REPORT*\n`;
    text += `📅 *Date:* ${todayStr}\n`;
    text += `🏫 *Class:* CLASS ${selectedClass ? selectedClass.toUpperCase() : ''}\n`;
    text += `👥 *Total Students:* ${classStus.length}\n`;
    text += `✅ *Present:* ${presentStudents.length} | ❌ *Absent:* ${absentStudents.length}`;
    if (leaveStudents.length > 0) text += ` | 🟡 *Leave:* ${leaveStudents.length}`;
    text += `\n----------------------------------------\n\n`;

    if (absentStudents.length > 0) {
      text += `❌ *ABSENT STUDENTS (${absentStudents.length}):*\n`;
      absentStudents.forEach((s, idx) => {
        text += `${idx + 1}. ${s.name} (Roll #${classStus.indexOf(s) + 1})\n`;
      });
      text += `\n`;
    } else {
      text += `🎉 *ALL STUDENTS PRESENT! (100% Attendance)*\n\n`;
    }

    if (presentStudents.length > 0) {
      text += `✅ *PRESENT STUDENTS (${presentStudents.length}):*\n`;
      presentStudents.forEach((s, idx) => {
        text += `${idx + 1}. ${s.name}\n`;
      });
    }

    return text;
  };

  const handlePostDiary = (e) => {
    e.preventDefault();
    if (!newDiaryContent.trim() || !selectedClass) return;
    const newEntry = {
      id: `diary-${Date.now()}`,
      class: selectedClass,
      date: new Date().toISOString().split('T')[0],
      title: `${newDiarySubject} Task`,
      subject: newDiarySubject,
      content: newDiaryContent.trim(),
      teacherName: 'Class Mentor'
    };
    setDiaries(prev => [newEntry, ...prev]);
    setNewDiaryContent('');
  };

  // Generate WhatsApp Report for a diary entry
  const generateWhatsAppReport = (entry) => {
    const todayStr = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    let text = `📓 *CALIPH SCHOOL - DAILY DIARY*\n`;
    text += `📅 *Date:* ${entry.date || todayStr}\n`;
    text += `🏫 *Class:* CLASS ${entry.class ? entry.class.toUpperCase() : ''}\n`;
    text += `📚 *Subject:* ${entry.subject}\n`;
    text += `👨‍🏫 *Teacher:* ${entry.teacherName}\n`;
    text += `----------------------------------------\n\n`;
    text += `📝 *${entry.title}*\n\n`;
    text += `${entry.content}\n\n`;
    text += `----------------------------------------\n`;
    text += `_Sent from Caliph Student Tracker_`;
    return text;
  };

  // Score transaction states - starts at 0 for every student selection session
  const [sessionStar, setSessionStar] = useState(0);
  const [sessionTally, setSessionTally] = useState(0);
  const [sessionStarReason, setSessionStarReason] = useState('');
  const [sessionTallyReason, setSessionTallyReason] = useState('');
  const [showTallySuggestions, setShowTallySuggestions] = useState(false);
  
  // Summary modal and sharing
  const [showSummaryModal, setShowSummaryModal] = useState(false);
  const [summaryText, setSummaryText] = useState('');
  
  // IR feature states
  const getFirstDayOfMonth = () => {
    const d = new Date();
    return new Date(d.getFullYear(), d.getMonth(), 1).toISOString().split('T')[0];
  };
  const getLastDayOfMonth = () => {
    const d = new Date();
    return new Date(d.getFullYear(), d.getMonth() + 1, 0).toISOString().split('T')[0];
  };
  const [showIRDateModal, setShowIRDateModal] = useState(false);
  const [irFromDate, setIrFromDate] = useState(getFirstDayOfMonth());
  const [irToDate, setIrToDate] = useState(getLastDayOfMonth());
  const [dateModalNextAction, setDateModalNextAction] = useState(null);
  const [showIRSelectModal, setShowIRSelectModal] = useState(false);
  const [showIRModal, setShowIRModal] = useState(false);
  const [selectedIRStudent, setSelectedIRStudent] = useState(null);
  const [irAssignedStudents, setIrAssignedStudents] = useState([]);
  const [irHistoryLogs, setIrHistoryLogs] = useState([]);

  useEffect(() => {
    if (showIRModal && selectedIRStudent) {
      const fetchHistory = async () => {
        if (selectedIRStudent === 'ALL') {
          const allLogs = [];
          for (const s of irAssignedStudents) {
            try {
              const res = await fetch(`/api/history/${s.id}`);
              if (res.ok) {
                const data = await res.json();
                allLogs.push(...data);
              }
            } catch (err) {}
          }
          allLogs.sort((a, b) => new Date(b.date) - new Date(a.date));
          setIrHistoryLogs(allLogs);
        } else {
          try {
            const res = await fetch(`/api/history/${selectedIRStudent.id}`);
            if (res.ok) {
              const data = await res.json();
              setIrHistoryLogs(data);
            }
          } catch (err) {
            console.error("Failed to fetch history:", err);
          }
        }
      };
      fetchHistory();
    }
  }, [showIRModal, selectedIRStudent, irAssignedStudents]);
  const [copyFeedback, setCopyFeedback] = useState('');
  const [saveStatus, setSaveStatus] = useState('');

  // Bulk text entry states
  const [showBulkModal, setShowBulkModal] = useState(false);
  const [bulkInputText, setBulkInputText] = useState('');

  // Load from DB on mount
  useEffect(() => {
    const loadFromDb = async () => {
      try {
        setDbLoading(true);
        const res = await fetch('/api/students');
        if (!res.ok) throw new Error('Database server returned an error');
        const data = await res.json();
        if (data && data.length > 0) {
          setStudents(data);
        } else {
          setStudents(INITIAL_STUDENTS);
        }
      } catch (err) {
        console.error('Error loading students from DB:', err);
        setDbError('Operating in offline mode with demo students.');
        // Fallback to localStorage
        const saved = localStorage.getItem('student_tracker_data');
        if (saved) {
          try {
            const parsed = JSON.parse(saved);
            if (parsed && parsed.length > 0) {
              setStudents(parsed);
            } else {
              setStudents(INITIAL_STUDENTS);
            }
          } catch (e) {
            console.error('Failed to parse cached student data:', e);
            setStudents(INITIAL_STUDENTS);
          }
        } else {
          setStudents(INITIAL_STUDENTS);
        }
      } finally {
        setDbLoading(false);
      }
    };
    loadFromDb();
  }, []);

  // Sync to localStorage as backup
  useEffect(() => {
    localStorage.setItem('student_tracker_data', JSON.stringify(students));
  }, [students]);

  // Debounced inline editor updates
  const updateTimeouts = React.useRef({});
  
  const debounceUpdateStudent = (studentId) => {
    setSaveStatus('Saving changes...');
    
    if (updateTimeouts.current[studentId]) {
      clearTimeout(updateTimeouts.current[studentId]);
    }
    
    updateTimeouts.current[studentId] = setTimeout(() => {
      setStudents(currentStudents => {
        const student = currentStudents.find(s => s.id === studentId);
        if (student) {
          fetch(`/api/students/${studentId}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(student)
          }).then(res => {
            if (res.ok) {
              setSaveStatus('Changes saved to DB');
              setTimeout(() => setSaveStatus(''), 2000);
            } else {
              setSaveStatus('Error saving changes');
            }
          }).catch(err => {
            console.error('API Error:', err);
            setSaveStatus('Connection error');
          });
        }
        return currentStudents;
      });
      delete updateTimeouts.current[studentId];
    }, 500);
  };

  // Clean up timeouts on unmount
  useEffect(() => {
    return () => {
      if (updateTimeouts.current) {
        Object.values(updateTimeouts.current).forEach(clearTimeout);
      }
    };
  }, []);

  // Reset session scores to zero when opening scoring board
  useEffect(() => {
    if (isScoring) {
      setSessionStar(0);
      setSessionTally(0);
      setSessionStarReason('');
      setSessionTallyReason('');
      setSaveStatus('');
    }
  }, [isScoring]);

  // Sync isScoring to false if selection becomes empty
  useEffect(() => {
    if (isScoring && selectedStudentIds.length === 0) {
      setIsScoring(false);
    }
  }, [selectedStudentIds, isScoring]);

  // Derived values
  const classStudents = selectedClass 
    ? students.filter(s => s.class === selectedClass) 
    : [];
  
  const filteredStudents = classStudents.filter((s, index) => {
    const query = searchQuery.toLowerCase().trim();
    const rollNumber = (index + 1).toString();
    return s.name.toLowerCase().includes(query) || rollNumber === query;
  });

  const PREDEFINED_TALLIES = [
    { reason: 'Misbehaviour in class', count: 3 },
    { reason: 'Without proper utensil in class', count: 3 },
    { reason: 'Sleeping on bench/desk', count: 3 },
    { reason: 'Late for Focus Hour (Night Study)', count: 3 },
    { reason: 'Not Doing Homework', count: 3 },
    { reason: 'Misplaced book', count: 3 },
    { reason: 'Leaving class without permission', count: 3 },
    { reason: 'Getting late to class after bell', count: 3 },
    { reason: 'Using podium without permission', count: 5 },
    { reason: 'Using mentors chair without permission', count: 5 },
    { reason: 'Group meetings without permission', count: 5 },
    { reason: 'practices without permission', count: 5 },
    { reason: 'Entering to mentors room without permission', count: 5 },
    { reason: "Using faculties' utensils", count: 5 },
    { reason: 'Eating in the class room without permission', count: 5 },
    { reason: 'Improper presentation of morning bliss', count: 5 },
    { reason: 'Playing while Prayer', count: 5 },
    { reason: 'talking while Prayer', count: 5 },
    { reason: 'Playing while Radio', count: 5 },
    { reason: 'Playing while Announcements', count: 5 },
    { reason: 'talking while Announcements', count: 5 },
    { reason: 'talking while Radio', count: 5 },
    { reason: 'Skipping programs and wandering in campus', count: 5 },
    { reason: 'Not presenting morning bliss', count: 10 },
    { reason: 'No Book for Pearl Hour', count: 10 }
  ];

  const filteredTallySuggestions = PREDEFINED_TALLIES.filter(item =>
    item.reason.toLowerCase().includes(sessionTallyReason.toLowerCase())
  );
  
  const totalStars = students.reduce((acc, curr) => acc + (Number(curr.star) || 0), 0);
  const totalTallies = students.reduce((acc, curr) => acc + (Number(curr.tally) || 0), 0);

  const selectedStudents = students.filter(s => selectedStudentIds.includes(s.id));

  // Score calculations
  const calculateTotalScore = (stars, tallies) => {
    return ((stars || 0) * 2) - ((tallies || 0) * 1);
  };

  const calculateGrade = (total) => {
    if (total >= 20) return 'A+';
    if (total >= 7) return 'A';
    if (total >= 0) return 'B';
    if (total >= -6) return 'C';
    if (total >= -20) return 'D';
    return 'E';
  };

  const calculateAttitudeGrade = (total) => {
    if (total >= -1.5) return 'A';
    if (total > -4) return 'B';
    if (total === -4) return 'C';
    if (total > -8) return 'D';
    return 'E';
  };

  const calculateNOGrade = (total) => {
    // Every 2 incidents drops the grade one letter.
    // 0 or -1 incidents → A
    // -2 or -3 incidents → B
    // -4 or -5 incidents → C
    // -6 or -7 incidents → D
    // -8 or below     → E
    if (total >= -1) return 'A';
    if (total >= -3) return 'B';
    if (total >= -5) return 'C';
    if (total >= -7) return 'D';
    return 'E';
  };

  const getNOIncidents = (s) => {
    if (!s) return 0;
    const inc = Number(s.neatAndOrderIncidents);
    if (!isNaN(inc) && inc > 0) {
      return inc;
    }
    const tally = Number(s.neatAndOrderTally) || 0;
    if (tally > 0) {
      return Math.max(1, Math.round(tally / 5));
    }
    return 0;
  };

  const getFineCount = (s) => (s.fineCount !== undefined && s.fineCount > 0) ? s.fineCount : (s.fine > 0 ? 1 : 0);

  // General field updater to support text descriptions
  const updateStudentField = (studentId, field, value) => {
    setStudents(prev => 
      prev.map(s => {
        if (s.id === studentId) {
          const updated = { ...s, [field]: value };
          if (field === 'neatAndOrderTally') {
            const val = Number(value) || 0;
            updated.neatAndOrderIncidents = val > 0 ? Math.max(1, Math.round(val / 5)) : 0;
          }
          return updated;
        }
        return s;
      })
    );
    debounceUpdateStudent(studentId);
  };

  const logHistory = (studentId, eventType, amount, reason) => {
    fetch('/api/history', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ student_id: studentId, event_type: eventType, amount, reason })
    }).catch(err => console.error("Error logging history:", err));
  };

  // Submit and log score session
  // Submit only star rating
  const handleSubmitStar = (e) => {
    if (e) e.preventDefault();
    if (selectedStudents.length === 0 || sessionStar <= 0) return;

    // Send to backend API
    const updatedStudentIds = [...selectedStudentIds];
    const amount = sessionStar;
    const reason = sessionStarReason.trim();

    updatedStudentIds.forEach(id => {
      logHistory(id, 'star', amount, reason);
    });

    fetch('/api/students/bulk-score', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ids: updatedStudentIds,
        type: 'star',
        amount: amount,
        reason: reason
      })
    }).catch(err => console.error("Error updating bulk star scores:", err));

    setStudents(prev => 
      prev.map(s => {
        if (selectedStudentIds.includes(s.id)) {
          return {
            ...s,
            star: s.star + sessionStar,
            starReason: sessionStarReason.trim() ? sessionStarReason.trim() : s.starReason
          };
        }
        return s;
      })
    );

    let lines = [];
    lines.push(`${sessionStar} star`);
    selectedStudents.forEach(student => {
      lines.push(student.name);
    });
    if (sessionStarReason.trim()) {
      lines.push('');
      let reason = sessionStarReason.trim();
      if (!reason.toLowerCase().startsWith('for ')) {
        reason = `for ${reason}`;
      }
      lines.push(reason);
    }

    const compiledText = lines.join('\n');
    setSummaryText(compiledText);
    setShowSummaryModal(true);

    // Reset star session inputs
    setSessionStar(0);
    setSessionStarReason('');
  };

  // Submit only tally rating
  const handleSubmitTally = (e) => {
    if (e) e.preventDefault();
    if (selectedStudents.length === 0 || sessionTally <= 0) return;

    // Send to backend API
    const updatedStudentIds = [...selectedStudentIds];
    const amount = sessionTally;
    const reason = sessionTallyReason.trim();

    updatedStudentIds.forEach(id => {
      logHistory(id, 'tally', amount, reason);
    });

    fetch('/api/students/bulk-score', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ids: updatedStudentIds,
        type: 'tally',
        amount: amount,
        reason: reason
      })
    }).catch(err => console.error("Error updating bulk tally scores:", err));

    setStudents(prev => 
      prev.map(s => {
        if (selectedStudentIds.includes(s.id)) {
          return {
            ...s,
            tally: s.tally + sessionTally,
            tallyReason: sessionTallyReason.trim() ? sessionTallyReason.trim() : s.tallyReason
          };
        }
        return s;
      })
    );

    let lines = [];
    lines.push(`${sessionTally} tally`);
    selectedStudents.forEach(student => {
      lines.push(student.name);
    });
    if (sessionTallyReason.trim()) {
      lines.push('');
      let reason = sessionTallyReason.trim();
      if (!reason.toLowerCase().startsWith('for ')) {
        reason = `for ${reason}`;
      }
      lines.push(reason);
    }

    const compiledText = lines.join('\n');
    setSummaryText(compiledText);
    setShowSummaryModal(true);

    // Reset tally session inputs
    setSessionTally(0);
    setSessionTallyReason('');
  };

  // Close summary and return automatically to the main screen
  const handleCloseSummaryAndGoHome = () => {
    setShowSummaryModal(false);
    setSelectedStudentIds([]);
    setIsScoring(false);
    setSelectedClass(null);
    setSearchQuery('');
  };

  const handleCopyText = () => {
    navigator.clipboard.writeText(summaryText).then(() => {
      setCopyFeedback('Copied!');
      setTimeout(() => setCopyFeedback(''), 2000);
    });
  };

  const handleAddStudent = (e) => {
    e.preventDefault();
    const targetClass = activeTab === 'admin' ? adminClass : selectedClass;
    if (!newStudentName.trim() || !targetClass) return;

    const newId = `${targetClass}-${Date.now()}`;
    const newStudent = {
      id: newId,
      name: newStudentName.trim(),
      class: targetClass,
      star: 0,
      tally: 0,
      starReason: '',
      tallyReason: ''
    };

    // Send to backend API
    fetch('/api/students', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(newStudent)
    }).catch(err => console.error("Error creating student:", err));

    setStudents(prev => [...prev, newStudent]);
    if (activeTab === 'scoring') {
      setSelectedStudentIds([newId]);
      setIsScoring(true);
    }
    setNewStudentName('');
    setShowAddModal(false);
  };

  const handleBulkSubmit = (e) => {
    if (e) e.preventDefault();
    if (!bulkInputText.trim()) return;

    const parsedEntries = parseBulkText(bulkInputText, students, selectedClass);
    if (parsedEntries.length === 0) return;

    const now = Date.now();

    // Clone parsed entries to avoid mutation issues
    const localEntries = parsedEntries.map(entry => ({
      ...entry,
      student: { ...entry.student }
    }));

    let updatedStudents = [...students];
    const studentsToUpsert = [];
    
    localEntries.forEach((entry, index) => {
      const amount = entry.amount;
      const type = entry.type;
      const reason = entry.reason;
      
      let targetStudent = null;
      if (entry.student.id) {
        targetStudent = updatedStudents.find(s => s.id === entry.student.id);
      } else {
        targetStudent = updatedStudents.find(s => 
          s.name.toLowerCase().trim() === entry.student.name.toLowerCase().trim() &&
          s.class === (selectedClass || 's2b')
        );
      }

      if (targetStudent) {
        const updated = {
          ...targetStudent,
          star: type === 'star' ? targetStudent.star + amount : targetStudent.star,
          tally: type === 'tally' ? targetStudent.tally + amount : targetStudent.tally,
          starReason: type === 'star' ? (reason || targetStudent.starReason) : targetStudent.starReason,
          tallyReason: type === 'tally' ? (reason || targetStudent.tallyReason) : targetStudent.tallyReason
        };
        updatedStudents = updatedStudents.map(s => s.id === targetStudent.id ? updated : s);
        studentsToUpsert.push(updated);
        entry.student.id = targetStudent.id;
        entry.student.class = targetStudent.class;
      } else {
        const targetClass = selectedClass || 's2b';
        const newId = `${targetClass}-${now}-${index}`;
        const newStudent = {
          id: newId,
          name: entry.student.name.trim(),
          class: targetClass,
          star: type === 'star' ? amount : 0,
          tally: type === 'tally' ? amount : 0,
          starReason: type === 'star' ? reason : '',
          tallyReason: type === 'tally' ? reason : ''
        };
        updatedStudents.push(newStudent);
        studentsToUpsert.push(newStudent);
        entry.student.id = newId;
        entry.student.class = targetClass;
      }
    });

    // Send the changed/added students to the bulk upsert endpoint
    fetch('/api/students/bulk-upsert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ students: studentsToUpsert })
    }).catch(err => console.error("Error bulk upserting students:", err));

    setStudents(updatedStudents);

    localEntries.forEach((entry, index) => {
      parsedEntries[index].student.id = entry.student.id;
      parsedEntries[index].student.class = entry.student.class;
    });

    // Build summary text grouping by rating type, amount, and reason to match the requested layout
    const groups = {};
    parsedEntries.forEach(entry => {
      const key = `${entry.amount}_${entry.type}_${entry.reason || ''}`;
      if (!groups[key]) {
        groups[key] = {
          amount: entry.amount,
          type: entry.type,
          reason: entry.reason,
          students: []
        };
      }
      groups[key].students.push(entry.student.name);
    });

    const summaryLines = [];
    Object.values(groups).forEach((group, index) => {
      const label = group.type === 'star' ? 'star' : 'tally';
      summaryLines.push(`${group.amount} ${label}`);
      group.students.forEach(name => {
        summaryLines.push(name);
      });
      if (group.reason) {
        summaryLines.push('');
        let reason = group.reason;
        if (!reason.toLowerCase().startsWith('for ')) {
          reason = `for ${reason}`;
        }
        summaryLines.push(reason);
      }
      if (index < Object.values(groups).length - 1) {
        summaryLines.push('');
      }
    });

    const compiledText = summaryLines.join('\n').trim();
    setSummaryText(compiledText);
    setShowSummaryModal(true);

    // Reset bulk input and close modal
    setBulkInputText('');
    setShowBulkModal(false);
  };

  const handleDeleteStudent = (studentId) => {
    if (confirm('Are you sure you want to remove this student?')) {
      fetch(`/api/students/${studentId}`, {
        method: 'DELETE'
      }).catch(err => console.error("Error deleting student:", err));

      setStudents(prev => prev.filter(s => s.id !== studentId));
      setSelectedStudentIds(prev => prev.filter(id => id !== studentId));
    }
  };

  const handleToggleDownloadClass = (clsName) => {
    setDownloadSelectedClasses(prev => {
      if (prev.includes(clsName)) {
        return prev.filter(c => c !== clsName);
      } else {
        return [...prev, clsName];
      }
    });
  };

  const handleToggleAllDownloadClasses = () => {
    if (downloadSelectedClasses.length === CLASSES.length) {
      setDownloadSelectedClasses([]);
    } else {
      setDownloadSelectedClasses(CLASSES);
    }
  };

  const handleToggleClearClass = (clsName) => {
    setClearSelectedClasses(prev => {
      if (prev.includes(clsName)) {
        return prev.filter(c => c !== clsName);
      } else {
        return [...prev, clsName];
      }
    });
  };

  const handleToggleAllClearClasses = () => {
    if (clearSelectedClasses.length === CLASSES.length) {
      setClearSelectedClasses([]);
    } else {
      setClearSelectedClasses(CLASSES);
    }
  };

  const handleClearClassesData = async () => {
    if (clearSelectedClasses.length === 0) return;
    
    const confirmMsg = `Are you sure you want to reset all stars, tallies, and reasons to zero/empty for selected classes: ${clearSelectedClasses.map(c => c.toUpperCase()).join(', ')}? This action CANNOT be undone!`;
    if (!confirm(confirmMsg)) return;
    
    try {
      const res = await fetch('/api/students/clear', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ classesToClear: clearSelectedClasses })
      });
      
      if (!res.ok) throw new Error('Failed to clear database');
      
      // Update local state to reset the scores and reasons of cleared classes
      setStudents(prev => prev.map(s => {
        if (clearSelectedClasses.includes(s.class)) {
          return {
            ...s,
            star: 0,
            tally: 0,
            starReason: '',
            tallyReason: '',
            diaryStar: 0,
            diaryTally: 0,
            neatAndOrderTally: 0,
            neatAndOrderReason: '',
            fine: 0,
            fineCount: 0,
            fineReason: '',
            sheetTally: 0
          };
        }
        return s;
      }));
      
      // Clear selected student IDs to avoid references to deleted students
      setSelectedStudentIds([]);
      
      setShowClearModal(false);
      alert('Selected classes cleared successfully.');
    } catch (err) {
      console.error(err);
      alert('Error clearing data. Please try again.');
    }
  };

  const handleDownloadExcel = async () => {
    if (downloadSelectedClasses.length === 0) return;
    setShowDownloadModal(false);

    let historyLogs = [];
    try {
      const res = await fetch('/api/history');
      if (res.ok) {
        historyLogs = await res.json();
      }
    } catch (e) {
      console.error(e);
    }

    const wb = XLSX.utils.book_new();

    downloadSelectedClasses.forEach(clsName => {
      const clsStudents = students.filter(s => s.class === clsName);
      if (clsStudents.length > 0) {
        const data = clsStudents.map(s => {
          const sLogs = historyLogs.filter(log => {
            if (log.student_id !== s.id) return false;
            const logDate = new Date(log.date).toISOString().split('T')[0];
            return (!irFromDate || logDate >= irFromDate) && (!irToDate || logDate <= irToDate);
          });

          let dynStar = 0, dynTally = 0, dynFine = 0, dynNo = 0, dynDiary = 0, dynSheet = 0;
          sLogs.forEach(log => {
             const type = log.event_type.toLowerCase();
             if (type === 'star') dynStar += Math.abs(log.amount);
             else if (type === 'tally') dynTally += Math.abs(log.amount);
             else if (type === 'spot fine') dynFine += 1;
             else if (type === 'n&o tally' || type === 'n&o') dynNo += Math.abs(log.amount);
             else if (type === 'diary tally') dynDiary += Math.abs(log.amount);
             else if (type.includes('sheet') || type === 'apology') dynSheet += Number(log.amount);
          });

          const finalTotal = calculateTotalScore(dynStar, dynTally);
          const attitudeTotal = (dynDiary * -0.5) + (dynFine * -1.5) + dynSheet;
          const noIncidents = getNOIncidents(s);
          const noTotal = -noIncidents;
          const noGrade = calculateNOGrade(noTotal);

          return {
            'Name': s.name,
            'Stars': dynStar,
            'Tallies': dynTally,
            'Total Score': finalTotal,
            'Grade': calculateGrade(finalTotal),
            'N&O Tally': dynNo || s.neatAndOrderTally || 0,
            'N&O Total': noTotal,
            'N&O Grade': noGrade,
            'Diary Tallies': dynDiary,
            'Sheets': dynSheet,
            'Fine': dynFine,
            'Total': attitudeTotal,
            'Attitude Grade': calculateAttitudeGrade(attitudeTotal)
          };
        });
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, clsName.toUpperCase());
      }
    });

    const filename = downloadSelectedClasses.length === 1 
      ? `Class_${downloadSelectedClasses[0].toUpperCase()}_Report.xlsx` 
      : 'Student_Tracker_Report.xlsx';

    XLSX.writeFile(wb, filename);
  };

  const handleImportExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });

      const classesToClear = new Set();
      const newStudentsToImport = [];
      const now = Date.now();

      wb.SheetNames.forEach((sheetName, sheetIdx) => {
        const sheetNameClean = sheetName.toLowerCase().trim();
        const matchedClass = CLASSES.find(c => c.toLowerCase() === sheetNameClean);
        const targetClass = matchedClass || (sheetNameClean !== '' ? sheetNameClean : (wb.SheetNames.length === 1 ? adminClass : null));

        if (targetClass) {
          if (!CLASSES.includes(targetClass)) {
            setCLASSES(prev => prev.includes(targetClass) ? prev : [...prev, targetClass]);
          }
          classesToClear.add(targetClass);
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
          
          if (rows.length > 0) {
            let nameColIdx = 0;
            let rolColIdx = -1;
            
            const headers = rows[0];
            let startRow = 0;
            
            if (headers) {
              const nameIdx = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'name');
              const rolIdx = headers.findIndex(h => h && (h.toString().toLowerCase().trim() === 'rol' || h.toString().toLowerCase().trim() === 'roll' || h.toString().toLowerCase().trim().includes('roll')));
              
              if (nameIdx !== -1) {
                nameColIdx = nameIdx;
                startRow = 1;
              }
              if (rolIdx !== -1) {
                rolColIdx = rolIdx;
                startRow = 1;
              }
            }

            for (let r = startRow; r < rows.length; r++) {
              const row = rows[r];
              const studentName = row[nameColIdx];
              const rollVal = rolColIdx !== -1 ? row[rolColIdx] : null;
              
              if (studentName && studentName.toString().trim()) {
                const uniqueRoll = rollVal ? rollVal.toString().trim() : (r + 1).toString();
                newStudentsToImport.push({
                  id: `${targetClass}-${uniqueRoll}-${now}`,
                  name: studentName.toString().trim(),
                  class: targetClass,
                  star: 0,
                  tally: 0,
                  starReason: '',
                  tallyReason: ''
                });
              }
            }
          }
        }
      });

      if (newStudentsToImport.length > 0) {
        const classesArray = Array.from(classesToClear);
        fetch('/api/students/import', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            classesToClear: classesArray,
            newStudents: newStudentsToImport
          })
        }).catch(err => console.error("Error importing roster:", err));

        setStudents(prev => {
          // Clear matching classes to overwrite their roster
          const filtered = prev.filter(s => !classesToClear.has(s.class));
          return [...filtered, ...newStudentsToImport];
        });
        alert(`Successfully imported ${newStudentsToImport.length} students! Existing rosters in class(es): ${classesArray.map(c => c.toUpperCase()).join(', ')} were replaced.`);
        setShowAddModal(false);
      } else {
        alert("No students could be parsed. Ensure sheet names match classes or contain 'name' and 'rol' columns.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  // Get initials for avatar display
  const getInitials = (name) => {
    return name
      .split(' ')
      .map(word => word[0])
      .join('')
      .toUpperCase()
      .substring(0, 2);
  };

  // View state mapping for slider translateX calculation
  let activeSlideIndex = 0; // Classes view
  if (selectedClass && !isScoring) {
    activeSlideIndex = 1; // Students view
  } else if (selectedClass && isScoring) {
    activeSlideIndex = 2; // Details view
  }

  // Generate translate class/style
  const getTranslateStyle = () => {
    return {
      transform: `translateX(-${activeSlideIndex * 33.3333}%)`
    };
  };

  // Admin spreadsheet list filtering
  const adminClassStudents = students.filter(s => s.class === adminClass);

  // --- Global Auth Gate ---
  if (authLoading) {
    return (
      <div className="fixed inset-0 flex items-center justify-center bg-[#0F172A]">
        <div className="w-10 h-10 border-4 border-blue-500 border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }
  if (!currentUser) {
    return <LoginPage onLogin={(user) => setCurrentUser(user)} />;
  }

  if (showSplash) {
    return (
      <div className="fixed inset-0 flex flex-col items-center justify-center bg-gradient-to-b from-[#0F172A] via-[#1A365D] to-[#0F172A] text-white z-[9999] overflow-hidden select-none animate-splash-exit font-sans">
        {/* Glow effects */}
        <div className="absolute top-1/4 left-1/4 w-[320px] h-[320px] bg-[#2A4365]/30 rounded-full blur-[100px] animate-pulse pointer-events-none" />
        <div className="absolute bottom-1/4 right-1/4 w-[320px] h-[320px] bg-[#0284C7]/20 rounded-full blur-[100px] animate-pulse pointer-events-none" style={{ animationDelay: '1s' }} />

        {/* Content container */}
        <div className="flex flex-col items-center max-w-sm px-6 text-center animate-fade-in">
          {/* Logo with pulse background */}
          <div className="relative mb-8 w-44 h-44 flex items-center justify-center">
            <div className="absolute inset-0 bg-white/10 rounded-full blur-2xl animate-pulse" />
            <img 
              src={logo} 
              alt="School Logo" 
              className="w-40 h-auto relative z-10 animate-logo-intro drop-shadow-[0_12px_24px_rgba(0,0,0,0.4)]"
            />
          </div>

          {/* Caliph School Text */}
          <h2 className="text-2xl font-extrabold tracking-wide text-white mb-1.5 animate-slide-up-1 font-sans">
            CALIPH SCHOOL
          </h2>
          <p className="text-xs font-semibold text-slate-300 tracking-widest uppercase mb-8 animate-slide-up-2">
            Student Management System
          </p>

          {/* Loading Progress Bar */}
          <div className="w-52 h-1.5 bg-white/15 rounded-full overflow-hidden border border-white/10 animate-slide-up-3">
            <div 
              className="h-full bg-gradient-to-r from-amber-400 via-sky-400 to-white rounded-full shadow-[0_0_12px_rgba(255,255,255,0.6)]" 
              style={{ 
                width: '0%', 
                animation: 'fillProgress 1.5s linear forwards' 
              }}
            />
          </div>
        </div>
      </div>
    );
  }


  return (
    <div className="w-full max-w-[480px] h-[100dvh] sm:h-[90vh] sm:min-h-[600px] flex flex-col bg-[#F8F9FA] border-none sm:border sm:border-slate-200 rounded-none sm:rounded-2xl shadow-none sm:shadow-[0_12px_36px_rgba(26,54,93,0.12)] overflow-hidden mx-auto sm:my-4 transition-all duration-300 font-sans">
      
      {/* Database status banner */}
      {dbLoading && (
        <div className="bg-[#1A365D]/10 border-b border-[#1A365D]/20 text-[#1A365D] text-[11px] font-bold text-center py-1.5 px-4 flex justify-center items-center gap-2 animate-pulse z-20 shrink-0">
          <div className="w-2 h-2 bg-[#1A365D] rounded-full"></div>
          Syncing with database...
        </div>
      )}
      {dbError && (
        <div className="bg-rose-50 border-b border-rose-200 text-rose-700 text-[11px] font-bold text-center py-1.5 px-4 z-20 shrink-0">
          ⚠️ {dbError}
        </div>
      )}
      {!dbLoading && !dbError && saveStatus && (
        <div className="bg-emerald-50 border-b border-emerald-200 text-emerald-800 text-[11px] font-bold text-center py-1.5 px-4 transition-all z-20 shrink-0">
          ✓ {saveStatus}
        </div>
      )}

      {/* App Header (Clean institutional Navy Blue header) */}
      <header className="flex justify-between items-center p-4 px-5 bg-[#1A365D] text-white shadow-header z-10 shrink-0">
        <div className="flex items-center gap-3">
          {/* Custom animated back buttons */}
          {activeTab === 'scoring' && activeSlideIndex === 1 && (
            <button 
              onClick={() => {
                setSelectedClass(null);
                setSelectedStudentIds([]);
                setSearchQuery('');
              }}
              className="p-1.5 rounded-full hover:bg-white/10 text-white/90 hover:text-white transition-all mr-0.5 active:scale-95 border border-white/20"
              title="Back to Classrooms"
            >
              <ChevronLeft className="w-5 h-5" />
            </button>
          )}
          {activeTab === 'scoring' && activeSlideIndex === 2 && (
            <button 
              onClick={() => {
                setIsScoring(false);
                setSaveStatus('');
              }}
              className="p-1.5 rounded-full hover:bg-white/10 text-white/90 hover:text-white transition-all mr-0.5 active:scale-95 border border-white/20"
              title="Back to Students"
            >
              <ChevronLeft className="w-5 h-5" />
            </button>
          )}
          {activeTab === 'mentor' && selectedMentorIndex !== null && (
            <button 
              onClick={() => {
                setSelectedMentorIndex(null);
                setMentorSubTab('diary');
                setShowAttendanceSummary(false);
              }}
              className="p-1.5 rounded-full hover:bg-white/10 text-white/90 hover:text-white transition-all mr-0.5 active:scale-95 border border-white/20"
              title="Back to Mentors"
            >
              <ChevronLeft className="w-5 h-5" />
            </button>
          )}

          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-white/15 border border-white/20 flex items-center justify-center shadow-inner">
              <School className="text-white w-5 h-5" />
            </div>
            <div>
              <h1 className="text-base font-extrabold tracking-tight text-white leading-tight font-sans">
                {activeTab === 'admin' && (isAdminAuthenticated ? 'Admin Management' : 'Score Sheet')}
                {activeTab === 'mentor' && 'Mentor Directory'}
                {activeTab === 'performance' && 'Performance Analytics'}
                {activeTab === 'scoring' && (
                  <>
                    {activeSlideIndex === 0 && 'School Management'}
                    {activeSlideIndex === 1 && `Class ${selectedClass.toUpperCase()}`}
                    {activeSlideIndex === 2 && 'Student Scoring'}
                  </>
                )}
              </h1>
              <p className="text-[10px] text-slate-200 font-semibold uppercase tracking-wider">
                {activeTab === 'admin' && (isAdminAuthenticated ? 'Master Score Sheet' : 'Score Sheet View')}
                {activeTab === 'mentor' && 'Teacher & Mentorship Directory'}
                {activeTab === 'performance' && 'Leaderboard & Class Insights'}
                {activeTab === 'scoring' && (
                  <>
                    {activeSlideIndex === 0 && 'Classrooms Overview'}
                    {activeSlideIndex === 1 && 'Student Roster'}
                    {activeSlideIndex === 2 && 'Evaluation Board'}
                  </>
                )}
              </p>
            </div>
          </div>
        </div>
        
      </header>

      {/* App Body - Renders either Scoring View (with slider) or Admin View */}
      <div className="flex-1 overflow-hidden relative flex flex-col bg-[#F8F9FA]">
        
        {activeTab === 'scoring' ? (
          /* SCORING TRACKER VIEW */
          <div className="flex-1 overflow-hidden relative">
            <div 
              className="flex w-[300%] h-full transition-transform duration-300 ease-out" 
              style={getTranslateStyle()}
            >
              
              {/* PAGE 1: Classrooms List View */}
              <div className="w-1/3 h-full flex flex-col p-5 overflow-y-auto gap-4">
                <div className="flex justify-between items-center text-slate-600">
                  <div className="flex items-center gap-2">
                    <GraduationCap className="text-[#1A365D] w-4.5 h-4.5" />
                    <span className="text-xs font-bold uppercase tracking-wider text-[#1A365D]">Active Classrooms</span>
                  </div>
                  <span className="text-[11px] font-semibold text-slate-500">{CLASSES.length} Classes</span>
                </div>
                
                <div className="grid grid-cols-1 gap-3">
                  {visibleClasses.map(clsName => {
                    const count = students.filter(s => s.class === clsName).length;
                    const studentLabel = count === 1 ? '1 Student' : `${count} Students`;
                    
                    return (
                      <button
                        key={clsName}
                        onClick={() => {
                          setSelectedClass(clsName);
                          setSelectedStudentIds([]);
                          setIsScoring(false);
                          setSearchQuery('');
                        }}
                        className="group flex justify-between items-center p-4 px-5 rounded-xl border border-slate-200 bg-white text-slate-800 shadow-sm hover:shadow-md hover:border-[#1A365D] hover:-translate-y-0.5 active:translate-y-0 transition-all duration-200"
                      >
                        <div className="flex items-center gap-3.5">
                          <div className="w-10 h-10 rounded-lg bg-slate-100 group-hover:bg-[#1A365D]/10 text-[#1A365D] flex items-center justify-center font-bold text-sm transition-colors">
                            <BookOpen className="w-5 h-5 text-[#1A365D]" />
                          </div>
                          <div className="text-left">
                            <span className="block text-base font-extrabold text-[#1A365D] tracking-wide uppercase">
                              {clsName.toUpperCase()}
                            </span>
                            <span className="text-[11px] font-medium text-slate-500">
                              Section {clsName.slice(0, 1).toUpperCase()}
                            </span>
                          </div>
                        </div>

                        <div className="flex items-center gap-2">
                          <span className="text-xs font-bold px-3 py-1 rounded-full bg-slate-100 border border-slate-200 text-[#1A365D] group-hover:bg-[#1A365D] group-hover:text-white transition-colors">
                            {studentLabel}
                          </span>
                          <ChevronRight className="w-4 h-4 text-slate-400 group-hover:text-[#1A365D] group-hover:translate-x-0.5 transition-all" />
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* PAGE 2: Students & Classroom Detail View with 3 Sub-Tabs */}
              <div className="w-1/3 h-full flex flex-col overflow-hidden bg-[#F8F9FA]">
                
                {/* Classroom Sub-Tabs Switch Bar removed as per request */}

                {/* SUB-TAB 1: SCORING DIRECTORY */}
                {roomTab === 'scoring' && (
                  <div className="flex-1 flex flex-col overflow-hidden">
                    {/* Search filter bar */}
                    <div className="p-3.5 border-b border-slate-200 bg-white flex gap-2">
                      <div className="relative flex-1">
                        <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 w-4 h-4" />
                        <input
                          type="text"
                          value={searchQuery}
                          onChange={e => setSearchQuery(e.target.value)}
                          placeholder="Search students by name or roll..."
                          className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all duration-200"
                        />
                      </div>
                      <button
                        onClick={() => setShowBulkModal(true)}
                        className="px-3 bg-[#1A365D]/5 border border-[#1A365D]/20 rounded-lg text-xs font-bold text-[#1A365D] hover:bg-[#1A365D] hover:text-white active:scale-95 transition-all flex items-center gap-1.5 shrink-0"
                        title="Bulk Text Entry"
                      >
                        <FileText className="w-3.5 h-3.5" />
                        <span>Import Text</span>
                      </button>
                    </div>

                    {/* Select All / Deselect All actions */}
                    {filteredStudents.length > 0 && (
                      <div className="px-4 py-2 bg-slate-100/70 border-b border-slate-200 flex justify-between items-center text-xs font-bold text-slate-600">
                        <span>{selectedStudentIds.length} of {filteredStudents.length} selected</span>
                        <div className="flex gap-3">
                          <button
                            onClick={() => {
                              const allIds = filteredStudents.map(s => s.id);
                              const allSelected = allIds.every(id => selectedStudentIds.includes(id));
                              if (allSelected) {
                                setSelectedStudentIds(prev => prev.filter(id => !allIds.includes(id)));
                              } else {
                                setSelectedStudentIds(prev => {
                                  const newIds = [...prev];
                                  allIds.forEach(id => {
                                    if (!newIds.includes(id)) newIds.push(id);
                                  });
                                  return newIds;
                                });
                              }
                            }}
                            className="text-[#1A365D] hover:underline transition-colors"
                          >
                            {filteredStudents.every(s => selectedStudentIds.includes(s.id)) ? 'Deselect All' : 'Select All'}
                          </button>
                          {selectedStudentIds.length > 0 && (
                            <button
                              onClick={() => setSelectedStudentIds([])}
                              className="text-rose-600 hover:underline transition-colors"
                            >
                              Clear
                            </button>
                          )}
                        </div>
                      </div>
                    )}

                    {/* Student card list */}
                    <div className="p-4 overflow-y-auto flex flex-col gap-2.5 flex-1">
                      {filteredStudents.length > 0 ? (
                        <>
                          {filteredStudents.map(student => {
                            const isSelected = selectedStudentIds.includes(student.id);
                            const rollNumber = classStudents.indexOf(student) + 1;
                            return (
                              <div
                                key={student.id}
                                onClick={() => {
                                  setSelectedStudentIds(prev => 
                                    prev.includes(student.id) 
                                      ? prev.filter(id => id !== student.id) 
                                      : [...prev, student.id]
                                  );
                                }}
                                className={`group flex justify-between items-center p-3.5 px-4 rounded-xl border transition-all duration-200 cursor-pointer ${
                                  isSelected 
                                    ? 'border-[#1A365D] bg-[#1A365D]/5 shadow-sm' 
                                    : 'border-slate-200 bg-white hover:bg-slate-50 hover:border-slate-300'
                                }`}
                              >
                                <div className="flex items-center gap-3 flex-1 min-w-0 mr-2">
                                  <div 
                                    className={`w-4 h-4 rounded border flex items-center justify-center transition-all shrink-0 ${
                                      isSelected 
                                        ? 'bg-[#1A365D] border-[#1A365D] text-white' 
                                        : 'border-slate-300 bg-white'
                                    }`}
                                  >
                                    {isSelected && <Check className="w-3 h-3 stroke-[3]" />}
                                  </div>
                                  <div className="flex-1 min-w-0">
                                    <span className="font-extrabold text-[#333333] text-sm block truncate group-hover:text-[#1A365D] transition-colors">
                                      {rollNumber} {student.name}
                                    </span>
                                    <div className="flex gap-2 text-[10px] font-bold mt-0.5">
                                      <span className="text-amber-600">⭐ {student.star} Stars</span>
                                      <span className="text-sky-600">🔢 {student.tally} Tallies</span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                        </>
                      ) : (
                        <div className="text-center py-8 text-slate-400 text-xs font-medium">No students found.</div>
                      )}
                      </div>
                      {/* Next Button Footer */}
                      {selectedStudentIds.length > 0 && (
                        <div className="p-4 bg-white border-t border-slate-200 shrink-0 shadow-[0_-4px_12px_rgba(0,0,0,0.02)] z-10">
                          <button
                            onClick={() => setIsScoring(true)}
                            className="w-full py-3.5 bg-[#1A365D] hover:bg-[#2A4365] text-white rounded-xl font-extrabold text-sm shadow-md transition-all flex items-center justify-center gap-2 active:scale-[0.98]"
                          >
                            Next
                            <ChevronRight className="w-5 h-5" />
                          </button>
                        </div>
                      )}
                    </div>
                  )}

                {/* SUB-TAB 2: NIGHT ATTENDANCE */}
                {roomTab === 'attendance' && (
                  <div className="flex-1 flex flex-col overflow-hidden p-4 gap-3">
                    {/* Attendance Header Summary Card */}
                    <div className="bg-white border border-slate-200 p-3.5 rounded-xl shadow-xs flex items-center justify-between shrink-0">
                      <div>
                        <h3 className="text-xs font-extrabold text-[#1A365D] uppercase tracking-wider">
                          Night Attendance — Class {selectedClass ? selectedClass.toUpperCase() : ''}
                        </h3>
                        <p className="text-[10px] text-slate-500 font-semibold mt-0.5">
                          {new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })} • Auto-marked Present
                        </p>
                      </div>
                      <div className="flex items-center gap-2">
                        <span className="text-[11px] font-extrabold px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-200">
                          {classStudents.filter(s => (attendanceMap[s.id] || 'present') === 'present').length} Present
                        </span>
                        <span className="text-[11px] font-extrabold px-2.5 py-1 rounded-full bg-rose-100 text-rose-800 border border-rose-200">
                          {classStudents.filter(s => (attendanceMap[s.id] || 'present') === 'absent').length} Absent
                        </span>
                      </div>
                    </div>

                    {/* Student Night Attendance List */}
                    <div className="flex-1 overflow-y-auto flex flex-col gap-2">
                      {classStudents.map((student, idx) => {
                        const status = attendanceMap[student.id] || 'present';
                        return (
                          <div
                            key={student.id}
                            className="bg-white border border-slate-200 p-3 px-4 rounded-xl shadow-xs flex items-center justify-between gap-3"
                          >
                            <div className="flex items-center gap-3">
                              <span className="w-6 h-6 rounded-md bg-slate-100 text-[#1A365D] text-xs font-extrabold flex items-center justify-center font-mono">
                                {idx + 1}
                              </span>
                              <div>
                                <h4 className="text-xs font-extrabold text-[#333333]">{student.name}</h4>
                                <span className="text-[10px] text-slate-400 font-semibold">Roll #{idx + 1}</span>
                              </div>
                            </div>

                            {/* Status Buttons Group */}
                            <div className="flex items-center gap-1">
                              <button
                                type="button"
                                onClick={() => setAttendanceMap(prev => ({ ...prev, [student.id]: 'present' }))}
                                className={`px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition-all ${
                                  status === 'present'
                                    ? 'bg-emerald-600 text-white shadow-xs'
                                    : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                                }`}
                              >
                                Present
                              </button>

                              <button
                                type="button"
                                onClick={() => setAttendanceMap(prev => ({ ...prev, [student.id]: 'absent' }))}
                                className={`px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition-all ${
                                  status === 'absent'
                                    ? 'bg-rose-600 text-white shadow-xs'
                                    : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                                }`}
                              >
                                Absent
                              </button>

                              <button
                                type="button"
                                onClick={() => setAttendanceMap(prev => ({ ...prev, [student.id]: 'leave' }))}
                                className={`px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition-all ${
                                  status === 'leave'
                                    ? 'bg-amber-500 text-white shadow-xs'
                                    : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                                }`}
                              >
                                Leave
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    {/* Submit & Share WhatsApp Action Button */}
                    <button
                      type="button"
                      onClick={() => {
                        const reportText = generateWhatsAppAttendanceReport();
                        setWhatsAppReportText(reportText);
                        setShowWhatsAppModal(true);
                      }}
                      className="w-full flex items-center justify-center gap-2 py-3 px-4 rounded-xl font-extrabold text-xs bg-emerald-600 hover:bg-emerald-700 text-white active:scale-[0.98] transition-all shadow-md shrink-0"
                    >
                      <Share2 className="w-4 h-4" />
                      Submit & Share Report via WhatsApp
                    </button>
                  </div>
                )}

                {/* SUB-TAB 3: DAILY DIARY */}
                {roomTab === 'diary' && (
                  <div className="flex-1 flex flex-col overflow-y-auto p-4 gap-4">
                    {/* Post Daily Diary Form */}
                    <form onSubmit={handlePostDiary} className="bg-white border border-slate-200 p-4 rounded-xl shadow-xs flex flex-col gap-3 shrink-0">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                        <span className="text-xs font-extrabold text-[#1A365D] uppercase tracking-wider flex items-center gap-2">
                          <FileText className="w-4 h-4 text-[#1A365D]" />
                          Post Class Daily Diary
                        </span>
                        <span className="text-[10px] font-bold text-slate-400">Class {selectedClass ? selectedClass.toUpperCase() : ''}</span>
                      </div>

                      <div>
                        <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Subject</label>
                        <select
                          value={newDiarySubject}
                          onChange={e => setNewDiarySubject(e.target.value)}
                          className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-800"
                        >
                          <option value="Science / Math">Science / Math</option>
                          <option value="English / Literature">English / Literature</option>
                          <option value="Islamic Studies">Islamic Studies</option>
                          <option value="Computer Science">Computer Science</option>
                          <option value="General Homework">General Homework</option>
                        </select>
                      </div>

                      <div>
                        <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Homework & Study Notes</label>
                        <textarea
                          rows="3"
                          required
                          placeholder="e.g. Complete Chapter 4 questions Q1-Q10. Prepare for tomorrow's test."
                          value={newDiaryContent}
                          onChange={e => setNewDiaryContent(e.target.value)}
                          className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-lg text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all font-medium resize-none"
                        ></textarea>
                      </div>

                      <button
                        type="submit"
                        className="w-full py-2.5 px-4 rounded-xl font-extrabold text-xs bg-[#1A365D] hover:bg-[#2A4365] text-white active:scale-[0.98] transition-all shadow-xs"
                      >
                        Post Daily Diary Entry
                      </button>
                    </form>

                    {/* Diary Entries Feed */}
                    <div className="flex flex-col gap-3">
                      <span className="text-xs font-extrabold text-[#1A365D] uppercase tracking-wider">
                        Class Diary Log ({diaries.filter(d => d.class === selectedClass).length})
                      </span>

                      {diaries.filter(d => d.class === selectedClass).length > 0 ? (
                        diaries.filter(d => d.class === selectedClass).map(entry => (
                          <div key={entry.id} className="bg-white border border-slate-200 p-4 rounded-xl shadow-xs flex flex-col gap-2.5">
                            <div className="flex justify-between items-start">
                              <div>
                                <h4 className="text-xs font-extrabold text-[#1A365D]">{entry.title}</h4>
                                <span className="text-[10px] font-semibold text-slate-400">Date: {entry.date} • Subject: {entry.subject}</span>
                              </div>
                              <span className="text-[10px] font-extrabold px-2 py-0.5 rounded-full bg-slate-100 text-slate-600">
                                {entry.teacherName}
                              </span>
                            </div>

                            <p className="text-xs text-slate-700 leading-relaxed font-medium bg-slate-50 p-2.5 rounded-lg border border-slate-100">
                              "{entry.content}"
                            </p>

                            <button
                              type="button"
                              onClick={() => {
                                const diaryReportText = generateWhatsAppReport(entry);
                                setWhatsAppReportText(diaryReportText);
                                setShowWhatsAppModal(true);
                              }}
                              className="self-end flex items-center gap-1.5 text-xs font-extrabold text-emerald-700 hover:text-emerald-800 bg-emerald-50 border border-emerald-200 px-3 py-1.5 rounded-lg transition-all"
                            >
                              <Share2 className="w-3.5 h-3.5" />
                              Share Diary via WhatsApp
                            </button>
                          </div>
                        ))
                      ) : (
                        <div className="text-xs text-slate-400 italic text-center py-6">
                          No daily diary entries posted yet for Class {selectedClass ? selectedClass.toUpperCase() : ''}.
                        </div>
                      )}
                    </div>
                  </div>
                )}
                </div>

              {/* PAGE 3: Student Details & Scoring View */}
              <div className="w-1/3 h-full flex flex-col p-5 overflow-y-auto gap-4 bg-[#F8F9FA]">
                {selectedStudents.length > 0 ? (
                  <form onSubmit={(e) => e.preventDefault()} className="flex flex-col gap-4">
                    {/* Profile panel */}
                    {selectedStudents.length === 1 ? (
                      <div className="flex items-center gap-4 pb-4 border-b border-slate-200 bg-white p-4 rounded-xl shadow-sm border">
                        <div className="w-12 h-12 rounded-xl bg-[#1A365D] text-white text-base font-extrabold flex items-center justify-center shadow-sm shrink-0">
                          {getInitials(selectedStudents[0].name)}
                        </div>
                        <div className="flex-1">
                          <h3 className="text-base font-extrabold text-[#333333] leading-tight">{selectedStudents[0].name}</h3>
                          <div className="flex justify-between items-center mt-1.5">
                            <span className="text-[10px] font-bold text-[#1A365D] bg-[#1A365D]/10 px-2.5 py-0.5 rounded-full uppercase tracking-wider">
                              Class {selectedStudents[0].class}
                            </span>
                            <span className="text-[11px] font-bold text-slate-600">
                              ⭐ {selectedStudents[0].star} | 🔢 {selectedStudents[0].tally}
                            </span>
                          </div>
                        </div>
                      </div>
                    ) : (
                      <div className="flex flex-col gap-2 pb-4 border-b border-slate-200 bg-white p-4 rounded-xl shadow-sm border">
                        <div className="flex items-center gap-3">
                          <div className="w-11 h-11 rounded-xl bg-[#1A365D] text-white text-sm font-extrabold flex items-center justify-center shadow-sm shrink-0">
                            {selectedStudents.length}
                          </div>
                          <div className="flex-1">
                            <h3 className="text-sm font-bold text-[#333333] leading-tight">
                              Evaluating {selectedStudents.length} Students
                            </h3>
                            <span className="text-[10px] font-bold text-[#1A365D] bg-[#1A365D]/10 px-2 py-0.5 rounded-full uppercase tracking-wider mt-1 inline-block">
                              Class {selectedClass?.toUpperCase()}
                            </span>
                          </div>
                        </div>
                        <div className="flex flex-wrap gap-1.5 mt-2.5 max-h-[85px] overflow-y-auto pt-2 border-t border-slate-100">
                          {selectedStudents.map(student => (
                            <span 
                              key={student.id} 
                              className="text-[11px] font-semibold bg-slate-100 border border-slate-200 pl-2.5 pr-1.5 py-0.5 rounded-full flex items-center gap-1.5 text-slate-700"
                            >
                              <span className="truncate max-w-[100px]">{student.name}</span>
                              <button 
                                type="button" 
                                onClick={() => {
                                  setSelectedStudentIds(prev => prev.filter(id => id !== student.id));
                                }}
                                className="w-4 h-4 rounded-full flex items-center justify-center bg-slate-200 hover:bg-rose-100 hover:text-rose-600 text-slate-500 transition-colors text-xs font-bold shrink-0"
                                title="Remove from selection"
                              >
                                <X className="w-2.5 h-2.5" />
                              </button>
                            </span>
                          ))}
                        </div>
                      </div>
                    )}

                    {/* Score Controls */}
                    <div className="flex flex-col gap-4">
                      {/* First Box: STAR */}
                      <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col gap-3 shadow-sm transition-all duration-200 focus-within:border-amber-500">
                        <div className="flex justify-between items-center">
                          <label className="text-xs font-extrabold text-[#333333] flex items-center gap-2 uppercase tracking-wide">
                            <Star className="w-4 h-4 text-amber-500 fill-amber-500/20" />
                            Star Points
                          </label>
                          <span className="text-[10px] text-amber-600 font-bold uppercase bg-amber-50 px-2 py-0.5 rounded-full border border-amber-200">+ Stars</span>
                        </div>

                        <div className="bg-amber-50/50 border border-amber-200 rounded-lg">
                          <input
                            type="number"
                            min="0"
                            placeholder="0"
                            value={sessionStar || ''}
                            onChange={e => setSessionStar(Math.max(0, parseInt(e.target.value) || 0))}
                            className="w-full bg-transparent text-center py-2.5 text-xl font-extrabold text-amber-800 focus:outline-none rounded-lg"
                          />
                        </div>

                        {/* Star Reason Typing Input */}
                        <div className="flex flex-col gap-1.5 mt-1 border-t border-slate-100 pt-3">
                          <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide flex items-center gap-1.5">
                            <FileText className="w-3 h-3 text-amber-600" />
                            Reason for Star
                          </label>
                          <input
                            type="text"
                            value={sessionStarReason}
                            onChange={e => setSessionStarReason(e.target.value)}
                            placeholder="e.g. Excellent class participation"
                            className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-amber-500 transition-colors"
                          />
                        </div>

                        {/* Submit Star Button */}
                        <button
                          type="button"
                          disabled={sessionStar === 0}
                          onClick={handleSubmitStar}
                          className="w-full flex items-center justify-center gap-1.5 py-2.5 px-4 rounded-lg font-bold text-xs bg-amber-500 text-white disabled:opacity-40 disabled:cursor-not-allowed hover:bg-amber-600 active:scale-[0.98] transition-all shadow-sm mt-1"
                        >
                          <Send className="w-3.5 h-3.5" />
                          Award Stars
                        </button>
                      </div>

                      {/* Second Box: TALLY */}
                      <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col gap-3 shadow-sm transition-all duration-200 focus-within:border-sky-500">
                        <div className="flex justify-between items-center">
                          <label className="text-xs font-extrabold text-[#333333] flex items-center gap-2 uppercase tracking-wide">
                            <Hash className="w-4 h-4 text-sky-600" />
                            Tally Record
                          </label>
                          <span className="text-[10px] text-sky-700 font-bold uppercase bg-sky-50 px-2 py-0.5 rounded-full border border-sky-200">Tallies</span>
                        </div>

                        <div className="bg-sky-50/50 border border-sky-200 rounded-lg">
                          <input
                            type="number"
                            min="0"
                            placeholder="0"
                            value={sessionTally || ''}
                            onChange={e => setSessionTally(Math.max(0, parseInt(e.target.value) || 0))}
                            className="w-full bg-transparent text-center py-2.5 text-xl font-extrabold text-sky-800 focus:outline-none rounded-lg"
                          />
                        </div>

                        {/* Tally Reason Typing Input */}
                        <div className="flex flex-col gap-1.5 mt-1 border-t border-slate-100 pt-3 relative">
                          <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide flex items-center gap-1.5">
                            <FileText className="w-3 h-3 text-sky-600" />
                            Reason for Tally
                          </label>
                          <input
                            type="text"
                            value={sessionTallyReason}
                            onChange={e => {
                              setSessionTallyReason(e.target.value);
                              setShowTallySuggestions(true);
                            }}
                            onFocus={() => setShowTallySuggestions(true)}
                            onBlur={() => {
                              setTimeout(() => setShowTallySuggestions(false), 200);
                            }}
                            placeholder="e.g. Missing notebook"
                            className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-sky-500 transition-colors"
                          />
                          {showTallySuggestions && sessionTallyReason && filteredTallySuggestions.length > 0 && (
                            <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-slate-200 rounded-xl shadow-lg max-h-48 overflow-y-auto z-50 divide-y divide-slate-100">
                              {filteredTallySuggestions.map((item, idx) => (
                                <button
                                  key={idx}
                                  type="button"
                                  onMouseDown={() => {
                                    setSessionTallyReason(item.reason);
                                    setSessionTally(item.count);
                                    setShowTallySuggestions(false);
                                  }}
                                  className="w-full text-left px-3 py-2 text-xs text-slate-700 hover:bg-[#1A365D]/5 hover:text-[#1A365D] transition-all flex justify-between items-center"
                                >
                                  <span className="font-semibold">{item.reason}</span>
                                  <span className="text-[10px] font-extrabold px-2 py-0.5 rounded-full bg-sky-100 border border-sky-200 text-sky-800">
                                    {item.count} tally
                                  </span>
                                </button>
                              ))}
                            </div>
                          )}
                        </div>

                        {/* Submit Tally Button */}
                        <button
                          type="button"
                          disabled={sessionTally === 0}
                          onClick={handleSubmitTally}
                          className="w-full flex items-center justify-center gap-1.5 py-2.5 px-4 rounded-lg font-bold text-xs bg-sky-600 text-white disabled:opacity-40 disabled:cursor-not-allowed hover:bg-sky-700 active:scale-[0.98] transition-all shadow-sm mt-1"
                        >
                          <Send className="w-3.5 h-3.5 text-white" />
                          Record Tally
                        </button>
                      </div>
                    </div>
                  </form>
                ) : (
                  <div className="flex flex-col items-center justify-center gap-3 text-center py-20 text-slate-400">
                    <Sliders className="w-10 h-10 text-slate-300" />
                    <p className="text-xs max-w-[200px] font-medium text-slate-500">No student selected.</p>
                  </div>
                )}
              </div>
            </div>
          </div>
        ) : activeTab === 'mentor' ? (
          /* MENTOR DASHBOARD & MENTEE ROOMS VIEW */
          <div className="flex-1 flex flex-col overflow-hidden bg-[#F8F9FA]">
            

            {/* Selected Mentor Content */}
            {(() => {
              if (selectedMentorIndex === null) {
                return (
                  <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                    {/* All Mentors Quick Directory */}
                    <div className="flex flex-col gap-3">
                      <div className="flex items-center justify-between mb-4 mt-2">
                        <span className="text-sm font-extrabold text-[#1A365D] uppercase tracking-wider flex items-center gap-2">
                          <UserCheck className="w-5 h-5 text-[#1A365D]" />
                          All Mentors ({mentors.length})
                        </span>
                        <div className="flex items-center gap-2">
                          <button 
                            onClick={() => setShowAddMentorModal(true)}
                            className="flex items-center gap-1.5 px-4 py-2 bg-[#1A365D] hover:bg-[#2A4365] text-white text-xs font-extrabold rounded-xl shadow-md hover:shadow-lg transition-all active:scale-[0.98]"
                          >
                            <Plus className="w-4 h-4" />
                            <span>Add Mentor</span>
                          </button>
                          <button 
                            onClick={() => setShowDeleteMentorModal(true)}
                            className="flex items-center gap-1.5 px-4 py-2 bg-rose-50 border border-rose-200 hover:bg-rose-600 hover:text-white hover:border-rose-600 text-rose-600 text-xs font-extrabold rounded-xl shadow-sm hover:shadow-md transition-all active:scale-[0.98]"
                          >
                            <Trash2 className="w-4 h-4" />
                            <span>Remove</span>
                          </button>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 gap-3">
                        {mentors.map((m, mIdx) => (
                          <button 
                            key={m.id}
                            onClick={() => setSelectedMentorIndex(mIdx)}
                            className="group relative bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex items-center justify-between transition-all duration-200 hover:shadow-md hover:border-blue-300 active:scale-[0.98] overflow-hidden"
                          >
                            <div className="flex items-center gap-4">
                              <div className="w-12 h-12 rounded-xl text-lg font-extrabold flex items-center justify-center transition-all duration-300 bg-blue-50 text-blue-700 group-hover:scale-110 group-hover:bg-blue-100 shadow-inner">
                                {getInitials(m.name)}
                              </div>
                              <div className="text-left flex flex-col justify-center">
                                <h4 className="text-[15px] font-extrabold text-slate-800 tracking-tight leading-tight">{m.name}</h4>
                                <span className="text-xs text-slate-500 font-semibold mt-0.5">{m.role || 'Mentor'}</span>
                              </div>
                            </div>
                            <div className="flex items-center gap-2 group-hover:-translate-x-1 transition-all duration-300">
                               {m.roomNumber && <span className="text-[10px] font-bold text-slate-500 bg-slate-100 px-2 py-1 rounded-md">{m.roomNumber.toUpperCase()}</span>}
                               <ChevronRight className="w-5 h-5 text-blue-500 opacity-0 group-hover:opacity-100 transition-all duration-300" />
                            </div>
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>
                );
              }

              const mentor = mentors[selectedMentorIndex];
              const assignedStudents = students.filter(s => (mentorStudents[mentor.name] || []).includes(s.id));

              return (
                <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                  
                  {/* Back button & Mentor Info Header */}
                  <div className="flex items-center gap-3 bg-white p-3 rounded-xl border border-slate-200 shadow-sm shrink-0">
                    <button 
                      onClick={() => setSelectedMentorIndex(null)}
                      className="w-8 h-8 rounded-lg bg-slate-100 flex items-center justify-center text-slate-600 hover:bg-[#1A365D] hover:text-white transition-colors"
                      title="Go Back"
                    >
                      <ChevronLeft className="w-5 h-5" />
                    </button>
                    <div className="w-10 h-10 rounded-xl bg-[#1A365D] text-white text-sm font-extrabold flex items-center justify-center shadow-sm shrink-0">
                      {getInitials(mentor.name)}
                    </div>
                    <div className="flex-1">
                      <h3 className="text-sm font-extrabold text-[#333333] leading-tight">{mentor.name}</h3>
                      <p className="text-[11px] text-slate-500 font-medium mt-0.5">{mentor.role} • Class {mentor.classAssigned.toUpperCase()}{mentor.roomNumber ? ` • ${mentor.roomNumber.toUpperCase()}` : ''}</p>
                    </div>
                    <button 
                      onClick={() => setShowMentorAddStudentModal(true)}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-[#1A365D] hover:bg-[#2A4365] text-white text-[11px] font-extrabold rounded-lg shadow-sm transition-colors active:scale-[0.98]"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Add Student</span>
                    </button>
                    <button 
                      onClick={() => setShowMentorRemoveStudentModal(true)}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-white text-[11px] font-extrabold rounded-lg shadow-sm transition-colors active:scale-[0.98]"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>Remove</span>
                    </button>
                  </div>

                  {/* Mentor Sub-Tabs (Diary / Night Attendance) */}
                  <div className="flex bg-slate-200/60 p-1 rounded-xl shrink-0">
                    <button
                      onClick={() => setMentorSubTab('diary')}
                      className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${
                        mentorSubTab === 'diary'
                          ? 'bg-white text-[#1A365D] shadow-sm'
                          : 'text-slate-500 hover:text-slate-700'
                      }`}
                    >
                      📖 Diary
                    </button>
                    <button
                      onClick={() => setMentorSubTab('attendance')}
                      className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${
                        mentorSubTab === 'attendance'
                          ? 'bg-white text-[#1A365D] shadow-sm'
                          : 'text-slate-500 hover:text-slate-700'
                      }`}
                    >
                      🌙 Night Attendance
                    </button>
                    <button
                      onClick={() => setMentorSubTab('summary')}
                      className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${
                        mentorSubTab === 'summary'
                          ? 'bg-white text-[#1A365D] shadow-sm'
                          : 'text-slate-500 hover:text-slate-700'
                      }`}
                    >
                      📊 Summary
                    </button>
                  </div>

                  {/* Mentor Add Student Modal */}
      {showMentorAddStudentModal && selectedMentorIndex !== null && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl border border-slate-200 overflow-hidden" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-2 text-[#1A365D]">
                <UserCheck className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider">Assign Mentee</h2>
              </div>
              <button 
                onClick={() => setShowMentorAddStudentModal(false)}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handleAddStudentToMentor} className="p-4 flex flex-col gap-4">
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  placeholder="Search students by name..."
                  value={mentorStudentSearch}
                  onChange={e => setMentorStudentSearch(e.target.value)}
                  className="flex-1 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold text-[#1A365D] focus:outline-none focus:border-[#1A365D] focus:ring-1 focus:ring-[#1A365D]"
                />
                <button
                  type="button"
                  onClick={() => {
                    const filteredIds = students.filter(s => s.name.toLowerCase().includes(mentorStudentSearch.toLowerCase())).map(s => s.id);
                    const allSelected = filteredIds.length > 0 && filteredIds.every(id => selectedStudentsForMentor.includes(id));
                    if (allSelected) {
                      setSelectedStudentsForMentor(prev => prev.filter(id => !filteredIds.includes(id)));
                    } else {
                      setSelectedStudentsForMentor(prev => [...new Set([...prev, ...filteredIds])]);
                    }
                  }}
                  className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg text-[10px] font-bold transition-colors whitespace-nowrap border border-slate-200"
                >
                  {students.filter(s => s.name.toLowerCase().includes(mentorStudentSearch.toLowerCase())).length > 0 && 
                   students.filter(s => s.name.toLowerCase().includes(mentorStudentSearch.toLowerCase())).every(s => selectedStudentsForMentor.includes(s.id))
                    ? 'Deselect All' : 'Select All'}
                </button>
              </div>
              <div className="flex flex-col gap-2 max-h-[400px] overflow-y-auto pr-1">
                {students
                  .filter(s => s.name.toLowerCase().includes(mentorStudentSearch.toLowerCase()))
                  .map(s => {
                  const isSelected = selectedStudentsForMentor.includes(s.id);
                  return (
                    <button 
                      key={s.id}
                      type="button"
                      onClick={() => setSelectedStudentsForMentor(prev => prev.includes(s.id) ? prev.filter(id => id !== s.id) : [...prev, s.id])}
                      className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${
                        isSelected 
                          ? 'border-[#1A365D] bg-[#1A365D]/5 ring-1 ring-[#1A365D]/20 shadow-sm' 
                          : 'border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50'
                      }`}
                    >
                      <div className="flex items-center gap-3">
                        <div className={`w-9 h-9 rounded-full flex items-center justify-center text-[11px] font-extrabold transition-colors ${
                          isSelected ? 'bg-[#1A365D] text-white' : 'bg-slate-100 text-[#1A365D]'
                        }`}>
                          {getInitials(s.name)}
                        </div>
                        <div>
                          <p className={`text-xs font-extrabold ${isSelected ? 'text-[#1A365D]' : 'text-[#333333]'}`}>
                            {s.name}
                          </p>
                          <p className={`text-[10px] font-medium ${isSelected ? 'text-[#1A365D]/70' : 'text-slate-500'}`}>
                            Class {s.class.toUpperCase()}
                          </p>
                        </div>
                      </div>
                      {isSelected && (
                        <div className="w-5 h-5 rounded-full bg-[#1A365D] text-white flex items-center justify-center">
                          <CheckCircle2 className="w-3.5 h-3.5" />
                        </div>
                      )}
                    </button>
                  );
                })}
              </div>

              <div className="flex justify-end gap-3 mt-2 border-t border-slate-100 pt-4">
                <button 
                  type="button"
                  onClick={() => {
                    setShowMentorAddStudentModal(false);
                    setSelectedStudentsForMentor([]);
                    setMentorStudentSearch('');
                  }}
                  className="px-4 py-2 text-slate-500 hover:bg-slate-100 rounded-lg font-bold text-xs transition-colors"
                >
                  Cancel
                </button>
                <button 
                  type="submit"
                  disabled={selectedStudentsForMentor.length === 0}
                  className="px-4 py-2 bg-[#1A365D] hover:bg-[#2A4365] text-white rounded-lg font-bold text-xs shadow-sm transition-colors disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-2"
                >
                  <UserCheck className="w-4 h-4" /> Assign {selectedStudentsForMentor.length > 0 ? `(${selectedStudentsForMentor.length})` : ''}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Mentor Remove Student Modal */}
      {showMentorRemoveStudentModal && selectedMentorIndex !== null && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl border border-slate-200 overflow-hidden" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-2 text-amber-600">
                <Trash2 className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider">Remove Mentee</h2>
              </div>
              <button 
                onClick={() => setShowMentorRemoveStudentModal(false)}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handleRemoveStudentFromMentor} className="p-4 flex flex-col gap-4">
              
              <div className="flex flex-col gap-2 max-h-[400px] overflow-y-auto pr-1">
                {assignedStudents.length > 0 ? assignedStudents.map(s => {
                  const isSelected = selectedStudentForMentorRemove === s.id;
                  return (
                    <button 
                      key={s.id}
                      type="button"
                      onClick={() => setSelectedStudentForMentorRemove(s.id)}
                      className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${
                        isSelected 
                          ? 'border-amber-500 bg-amber-500/5 ring-1 ring-amber-500/20 shadow-sm' 
                          : 'border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50'
                      }`}
                    >
                      <div className="flex items-center gap-3">
                        <div className={`w-9 h-9 rounded-full flex items-center justify-center text-[11px] font-extrabold transition-colors ${
                          isSelected ? 'bg-amber-500 text-white' : 'bg-slate-100 text-amber-600'
                        }`}>
                          {getInitials(s.name)}
                        </div>
                        <div>
                          <p className={`text-xs font-extrabold ${isSelected ? 'text-amber-700' : 'text-[#333333]'}`}>
                            {s.name}
                          </p>
                          <p className={`text-[10px] font-medium ${isSelected ? 'text-amber-600/70' : 'text-slate-500'}`}>
                            Class {s.class.toUpperCase()}
                          </p>
                        </div>
                      </div>
                      {isSelected && (
                        <div className="w-5 h-5 rounded-full bg-amber-500 text-white flex items-center justify-center">
                          <CheckCircle2 className="w-3.5 h-3.5" />
                        </div>
                      )}
                    </button>
                  );
                }) : (
                  <p className="text-center text-slate-500 text-xs py-4">No students to remove.</p>
                )}
              </div>

              <div className="flex justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setShowMentorRemoveStudentModal(false)}
                  className="px-4 py-2 rounded-xl text-xs font-bold text-slate-500 bg-slate-100 hover:bg-slate-200 transition-colors"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={!selectedStudentForMentorRemove}
                  className="px-4 py-2 rounded-xl text-xs font-bold text-white bg-amber-500 hover:bg-amber-600 transition-colors shadow-md disabled:opacity-50 flex items-center gap-1.5"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  Remove
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

                  {/* Sub-Tab Content Area */}
                  <div className="flex-1 bg-white border border-slate-200 rounded-xl p-6 flex flex-col shadow-sm">
                    {mentorSubTab === 'diary' ? (
                      <div className="flex flex-col items-center gap-3 animate-fade-in w-full h-full relative overflow-hidden">
                        <div className="flex items-center gap-2 mb-2 shrink-0">
                          <span className="text-xl">📖</span>
                          <h3 className="text-sm font-extrabold text-[#1A365D]">Mentor Diary</h3>
                        </div>
                        
                        <div className="w-full flex-1 flex flex-col gap-2 max-w-xl mx-auto overflow-hidden">
                          <div className="flex items-center gap-4 bg-slate-50 p-3 rounded-xl border border-slate-200 shrink-0 mb-1">
                            <div className="flex-1 flex flex-col gap-1">
                              <label className="text-[10px] font-extrabold text-slate-500 uppercase">Start Date</label>
                              <input 
                                type="date" 
                                value={diaryStartDate}
                                onChange={(e) => setDiaryStartDate(e.target.value)}
                                className="w-full p-2 text-xs font-semibold bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-[#1A365D]"
                              />
                            </div>
                            <div className="flex-1 flex flex-col gap-1">
                              <label className="text-[10px] font-extrabold text-slate-500 uppercase">End Date</label>
                              <input 
                                type="date" 
                                value={diaryEndDate}
                                onChange={(e) => setDiaryEndDate(e.target.value)}
                                className="w-full p-2 text-xs font-semibold bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-[#1A365D]"
                              />
                            </div>
                          </div>

                          {showDiarySummary ? (
                            <div className="flex flex-col gap-3 animate-fade-in text-left overflow-y-auto">
                              <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
                                <h4 className="text-xs font-extrabold text-[#1A365D] uppercase tracking-wider mb-2">Diary Summary</h4>
                                {(() => {
                                  const formatStr = (ds) => {
                                    if(!ds) return '';
                                    const p = ds.split('-');
                                    return p.length === 3 ? `${p[2]}/${p[1]}/${p[0]}` : ds;
                                  };
                                  const start = formatStr(diaryStartDate);
                                  const end = formatStr(diaryEndDate);
                                  const periodStr = `${start} to ${end}`;

                                  let rawText = `mentor name ${mentor.name}\nperiod ${periodStr}\n${mentor.classAssigned.toUpperCase()}\nstudents name\n\n`;

                                  assignedStudents.forEach(s => {
                                    const record = diaryRecords[s.id] || { status: 'none', days: '' };
                                    if (record.status === 'written') {
                                      rawText += `${s.name} ✅\n`;
                                    } else if (record.status === 'not_written') {
                                      const daysStr = (record.days || '').replace(/[^0-9]/g, '');
                                      const numDays = parseInt(daysStr) || 0;
                                      const tallies = (numDays * (numDays + 1)) / 2;
                                      rawText += `${s.name}  (not submitted ${record.days || numDays + ' day'})  ${tallies} tally\n`;
                                    } else {
                                      rawText += `${s.name} (not submitted)\n`;
                                    }
                                  });
                                  
                                  const whatsappText = encodeURIComponent(rawText.trim());

                                  return (
                                    <div className="flex flex-col gap-3">
                                      <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 font-mono text-xs text-slate-700 whitespace-pre-wrap leading-relaxed shadow-inner">
                                        {rawText.trim()}
                                      </div>
                                      
                                      <div className="flex flex-col gap-2 mt-2 pt-2 border-t border-slate-100">
                                        <div className="flex gap-2">
                                          <button onClick={() => setShowDiarySummary(false)} className="flex-1 py-2 rounded-xl text-xs font-bold text-slate-500 bg-slate-100 hover:bg-slate-200 transition-colors">Edit</button>
                                          <a 
                                            href={`https://wa.me/?text=${whatsappText}`}
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            className="flex-1 py-2 rounded-xl text-xs font-bold text-white bg-green-500 hover:bg-green-600 shadow-sm transition-colors text-center flex items-center justify-center gap-2"
                                          >
                                            Send to WhatsApp
                                          </a>
                                        </div>
                                        <button 
                                          onClick={() => {
                                            setShowDiarySummary(false);
                                            setMentorSubTab(null);
                                          }} 
                                          className="w-full py-2.5 rounded-xl text-xs font-bold text-white bg-[#1A365D] hover:bg-[#2A4365] shadow-sm transition-colors mt-1"
                                        >
                                          Done, Return to Overview
                                        </button>
                                      </div>
                                    </div>
                                  );
                                })()}
                              </div>
                            </div>
                          ) : (
                            <div className="flex-1 flex flex-col relative overflow-hidden">
                              <div className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider bg-slate-50 py-1.5 px-3 rounded-md border border-slate-200 mb-1 flex justify-between items-center shrink-0">
                                <span>Assigned Students</span>
                                <span>{assignedStudents.length} Students</span>
                              </div>
                              
                              <div className="overflow-y-auto flex-1 pr-1 flex flex-col gap-3 pb-20">
                                {assignedStudents.length > 0 ? assignedStudents.map(student => {
                                  const record = diaryRecords[student.id] || { status: 'none', days: '' };
                                  const currentStatus = record.status;
                                  return (
                                    <div key={student.id} className="flex flex-col bg-white border border-slate-200 p-3 rounded-xl hover:shadow-xs hover:border-[#1A365D]/30 transition-all shrink-0">
                                      <div className="flex justify-between items-center">
                                        <div className="flex items-center gap-3">
                                          <div className="w-8 h-8 rounded-full bg-[#1A365D]/10 flex items-center justify-center text-[10px] font-extrabold text-[#1A365D]">
                                            {getInitials(student.name)}
                                          </div>
                                          <div className="text-left">
                                            <span className="text-xs font-bold text-[#1A365D] flex items-center gap-1.5">
                                              {student.ineligible && <div className="w-2 h-2 rounded-full bg-rose-600 animate-pulse" title={`Ineligible: ${student.ineligibleReason || 'No reason'}`} />}
                                              {student.name}
                                            </span>
                                            <span className="text-[9px] text-slate-400 font-medium">Class {student.class.toUpperCase()}</span>
                                          </div>
                                        </div>
                                        <div className="flex items-center gap-1">
                                          <button 
                                            onClick={() => handleDiaryChange(student.id, 'status', 'written')}
                                            className={`text-[10px] px-2.5 py-1.5 rounded-lg shadow-sm font-bold transition-colors ${currentStatus === 'written' ? 'bg-[#1A365D] text-white border-[#1A365D]' : 'bg-slate-50 border border-slate-200 text-slate-500 hover:bg-[#1A365D]/10 hover:text-[#1A365D]'}`}
                                          >
                                            Written
                                          </button>
                                          <button 
                                            onClick={() => handleDiaryChange(student.id, 'status', 'not_written')}
                                            className={`text-[10px] px-2.5 py-1.5 rounded-lg shadow-sm font-bold transition-colors ${currentStatus === 'not_written' ? 'bg-amber-500 text-white border-amber-500' : 'bg-slate-50 border border-slate-200 text-slate-500 hover:bg-amber-50 hover:text-amber-600'}`}
                                          >
                                            Not Written
                                          </button>
                                        </div>
                                      </div>
                                      
                                      {(currentStatus === 'written' || currentStatus === 'not_written') && (
                                        <div className="mt-3 pt-3 border-t border-slate-100 animate-fade-in">
                                          <input
                                            type="text"
                                            placeholder={currentStatus === 'not_written' ? "How many days? (e.g. 1, 2...)" : "Add an optional note (e.g. page numbers, topics)..."}
                                            value={record.days || ''}
                                            onChange={(e) => handleDiaryChange(student.id, 'days', e.target.value)}
                                            className={`w-full text-xs py-1.5 px-3 rounded-lg focus:outline-none focus:ring-2 font-medium transition-colors ${
                                              currentStatus === 'not_written' 
                                                ? 'bg-amber-50 border border-amber-200 focus:border-amber-400 focus:ring-amber-400/20 text-amber-900 placeholder-amber-400/70' 
                                                : 'bg-slate-50 border border-slate-200 focus:border-[#1A365D] focus:ring-[#1A365D]/20 text-slate-700 placeholder-slate-400'
                                            }`}
                                          />
                                        </div>
                                      )}
                                    </div>
                                  );
                                }) : (
                                  <div className="text-center py-10 flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-xl">
                                    <span className="text-3xl opacity-50 mb-2">👥</span>
                                    <span className="text-xs text-slate-400 font-medium">No students assigned yet.</span>
                                    <span className="text-[10px] text-slate-400 mt-1">Click "Add Student" above to assign them.</span>
                                  </div>
                                )}
                              </div>
                              
                              {assignedStudents.length > 0 && (
                                <div className="absolute bottom-0 left-0 right-0 pt-2 bg-white flex justify-center border-t border-slate-100 shrink-0">
                                  <button
                                    onClick={() => {
                                      // Sync tallies with global state based on changes
                                      assignedStudents.forEach(s => {
                                        const record = diaryRecords[s.id];
                                        if (record) {
                                          let tallies = 0;
                                          if (record.status === 'not_written') {
                                            const daysStr = (record.days || '').replace(/[^0-9]/g, '');
                                            const numDays = parseInt(daysStr) || 0;
                                            tallies = (numDays * (numDays + 1)) / 2;
                                          }
                                          const previouslyApplied = record.appliedTallies || 0;
                                          
                                          const difference = tallies - previouslyApplied;
                                          if (difference !== 0) {
                                            const currentTally = Number(s.diaryTally) || 0;
                                            updateStudentField(s.id, 'diaryTally', currentTally + difference);
                                          }
                                        }
                                      });
                                      
                                      // Mark records with appliedTallies to prevent double-counting
                                      setDiaryRecords(prev => {
                                        const next = { ...prev };
                                        assignedStudents.forEach(s => {
                                          const record = next[s.id];
                                          if (record) {
                                            let tallies = 0;
                                            if (record.status === 'not_written') {
                                              const daysStr = (record.days || '').replace(/[^0-9]/g, '');
                                              const numDays = parseInt(daysStr) || 0;
                                              tallies = (numDays * (numDays + 1)) / 2;
                                            }
                                            next[s.id] = { ...record, appliedTallies: tallies, submitted: true };
                                          }
                                        });
                                        return next;
                                      });

                                      setShowDiarySummary(true);
                                    }}
                                    className="w-full py-3 rounded-xl text-xs font-bold text-white bg-[#1A365D] hover:bg-[#2A4365] transition-colors shadow-md"
                                  >
                                    Submit Diary
                                  </button>
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      </div>
                    ) : mentorSubTab === 'attendance' ? (
                      <div className="flex flex-col items-center gap-3 animate-fade-in w-full h-full relative overflow-hidden">
                        <div className="flex items-center gap-2 mb-2 shrink-0">
                          <span className="text-xl">🌙</span>
                          <h3 className="text-sm font-extrabold text-[#1A365D]">Night Attendance</h3>
                        </div>
                        
                        <div className="w-full flex-1 flex flex-col gap-2 max-w-xl mx-auto overflow-hidden">
                          {showAttendanceSummary ? (
                            <div className="flex flex-col gap-3 animate-fade-in text-left overflow-y-auto">
                              <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
                                <h4 className="text-xs font-extrabold text-[#1A365D] uppercase tracking-wider mb-2">Attendance Summary</h4>
                                {(() => {
                                  const summaryList = assignedStudents.map(s => {
                                    const rec = attendanceRecords[s.id] || { status: 'present', reason: '' };
                                    return { ...s, status: rec.status || 'present', reason: rec.reason || '' };
                                  });
                                  const presentCount = summaryList.filter(s => s.status === 'present').length;
                                  const nonPresent = summaryList.filter(s => s.status !== 'present');
                                  
                                  const today = new Date();
                                  const dateStr = today.toLocaleDateString('en-GB'); // "dd/mm/yyyy"
                                  const timeStr = today.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}); 
                                  
                                  const absentees = summaryList.filter(s => s.status === 'absent');
                                  const lateStudents = summaryList.filter(s => s.status === 'late');

                                  let rawText = `date ${dateStr}\ntime ${timeStr}\nroom ${mentor.roomNumber ? mentor.roomNumber.toUpperCase() : 'N/A'}\nmentor name ${mentor.name}\n`;
                                  
                                  if (absentees.length === 0 && lateStudents.length === 0) {
                                    rawText += `ALL PRESENT\n`;
                                  } else {
                                    if (absentees.length > 0) {
                                      rawText += `ABSENT\n`;
                                      rawText += absentees.map(s => {
                                        const hasNoReason = !s.reason || s.reason.trim() === '';
                                        return hasNoReason ? `${s.name}  10 TALLY` : `${s.name} (${s.reason})`;
                                      }).join('\n') + '\n\n';
                                    }
                                    
                                    if (lateStudents.length > 0) {
                                      rawText += `LATE\n`;
                                      rawText += lateStudents.map(s => {
                                        const hasNoReason = !s.reason || s.reason.trim() === '';
                                        return hasNoReason ? `${s.name}   3 TALLY` : `${s.name} (${s.reason})`;
                                      }).join('\n') + '\n';
                                    }
                                  }
                                  
                                  const whatsappText = encodeURIComponent(rawText.trim());

                                  return (
                                    <div className="flex flex-col gap-3">
                                      <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 font-mono text-xs text-slate-700 whitespace-pre-wrap leading-relaxed shadow-inner">
                                        {rawText.trim()}
                                      </div>
                                      
                                      <div className="flex flex-col gap-2 mt-2 pt-2 border-t border-slate-100">
                                        <div className="flex gap-2">
                                          <button onClick={() => setShowAttendanceSummary(false)} className="flex-1 py-2 rounded-xl text-xs font-bold text-slate-500 bg-slate-100 hover:bg-slate-200 transition-colors">Edit</button>
                                          <a 
                                            href={`https://wa.me/?text=${whatsappText}`}
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            className="flex-1 py-2 rounded-xl text-xs font-bold text-white bg-green-500 hover:bg-green-600 shadow-sm transition-colors text-center flex items-center justify-center gap-2"
                                          >
                                            Send to WhatsApp
                                          </a>
                                        </div>
                                        <button 
                                          onClick={() => {
                                            setShowAttendanceSummary(false);
                                            setMentorSubTab(null);
                                          }} 
                                          className="w-full py-2.5 rounded-xl text-xs font-bold text-white bg-[#1A365D] hover:bg-[#2A4365] shadow-sm transition-colors mt-1"
                                        >
                                          Done, Return to Overview
                                        </button>
                                      </div>
                                    </div>
                                  );
                                })()}
                              </div>
                            </div>
                          ) : (
                            <div className="flex-1 flex flex-col relative overflow-hidden">
                              <div className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider bg-slate-50 py-1.5 px-3 rounded-md border border-slate-200 mb-1 flex justify-between items-center shrink-0">
                                <span>Assigned Students</span>
                                <span>{assignedStudents.length} Students</span>
                              </div>
                              
                              <div className="overflow-y-auto flex-1 pr-1 flex flex-col gap-3 pb-20">
                                {assignedStudents.length > 0 ? assignedStudents.map(student => {
                                  const record = attendanceRecords[student.id] || { status: 'present', reason: '' };
                                  const currentStatus = record.status || 'present';
                                  
                                  return (
                                    <div key={student.id} className="flex flex-col bg-white border border-slate-200 p-3 rounded-xl hover:shadow-xs hover:border-[#1A365D]/30 transition-all shrink-0">
                                      <div className="flex justify-between items-center">
                                        <div className="flex items-center gap-3">
                                          <div className="w-8 h-8 rounded-full bg-slate-100 flex items-center justify-center text-[10px] font-extrabold text-slate-600">
                                            {getInitials(student.name)}
                                          </div>
                                          <div className="text-left">
                                            <span className="text-xs font-bold text-[#1A365D] flex items-center gap-1.5">
                                              {student.ineligible && <div className="w-2 h-2 rounded-full bg-rose-600 animate-pulse" title={`Ineligible: ${student.ineligibleReason || 'No reason'}`} />}
                                              {student.name}
                                            </span>
                                            <span className="text-[9px] text-slate-400 font-medium">Class {student.class.toUpperCase()}</span>
                                          </div>
                                        </div>
                                        <div className="flex items-center gap-1">
                                          <button 
                                            onClick={() => handleAttendanceChange(student.id, 'status', 'present')}
                                            className={`text-[10px] px-2 py-1.5 rounded-lg shadow-sm font-bold transition-colors ${currentStatus === 'present' ? 'bg-emerald-500 text-white border-emerald-500' : 'bg-slate-50 border border-slate-200 text-slate-500 hover:bg-emerald-50 hover:text-emerald-600'}`}
                                          >
                                            Present
                                          </button>
                                          <button 
                                            onClick={() => handleAttendanceChange(student.id, 'status', 'late')}
                                            className={`text-[10px] px-2 py-1.5 rounded-lg shadow-sm font-bold transition-colors ${currentStatus === 'late' ? 'bg-amber-500 text-white border-amber-500' : 'bg-slate-50 border border-slate-200 text-slate-500 hover:bg-amber-50 hover:text-amber-600'}`}
                                          >
                                            Late
                                          </button>
                                          <button 
                                            onClick={() => handleAttendanceChange(student.id, 'status', 'absent')}
                                            className={`text-[10px] px-2 py-1.5 rounded-lg shadow-sm font-bold transition-colors ${currentStatus === 'absent' ? 'bg-rose-500 text-white border-rose-500' : 'bg-slate-50 border border-slate-200 text-slate-500 hover:bg-rose-50 hover:text-rose-600'}`}
                                          >
                                            Absent
                                          </button>
                                        </div>
                                      </div>
                                      
                                      {currentStatus !== 'present' && (
                                        <div className="mt-3 pt-3 border-t border-slate-100 animate-fade-in">
                                          <input
                                            type="text"
                                            placeholder={`Reason for being ${currentStatus}... (Optional)`}
                                            value={record.reason || ''}
                                            onChange={(e) => handleAttendanceChange(student.id, 'reason', e.target.value)}
                                            className="w-full text-xs py-1.5 px-3 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:border-[#1A365D] focus:ring-2 focus:ring-[#1A365D]/10"
                                          />
                                        </div>
                                      )}
                                    </div>
                                  );
                                }) : (
                                  <div className="text-center py-10 flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-xl">
                                    <span className="text-3xl opacity-50 mb-2">🛏️</span>
                                    <span className="text-xs text-slate-400 font-medium">No students assigned for attendance.</span>
                                  </div>
                                )}
                              </div>
                              
                              {assignedStudents.length > 0 && (
                                <div className="absolute bottom-0 left-0 right-0 pt-2 bg-white flex justify-center border-t border-slate-100 shrink-0">
                                  <button
                                    onClick={() => {
                                      setShowAttendanceSummary(true);
                                      
                                      const updatedStudents = [...students];
                                      const studentsToUpsert = [];
                                      const updatedRecords = { ...attendanceRecords };
                                      let hasChanges = false;
                                      
                                      assignedStudents.forEach(s => {
                                        const rec = attendanceRecords[s.id] || { status: 'present', reason: '' };
                                        if (rec.tallied) return; // Prevent double tallying in the same session
                                        
                                        const hasNoReason = !rec.reason || rec.reason.trim() === '';
                                        
                                        if (hasNoReason) {
                                          if (rec.status === 'absent') {
                                            const studentIndex = updatedStudents.findIndex(st => st.id === s.id);
                                            if (studentIndex !== -1) {
                                              const st = { ...updatedStudents[studentIndex] };
                                              st.tally = (st.tally || 0) + 10;
                                              st.tallyReason = 'Night Attendance (Absent, No Reason)';
                                              updatedStudents[studentIndex] = st;
                                              studentsToUpsert.push(st);
                                              logHistory(s.id, 'tally', 10, 'Night Attendance (Absent, No Reason)');
                                            }
                                          } else if (rec.status === 'late') {
                                            const studentIndex = updatedStudents.findIndex(st => st.id === s.id);
                                            if (studentIndex !== -1) {
                                              const st = { ...updatedStudents[studentIndex] };
                                              st.tally = (st.tally || 0) + 3;
                                              st.tallyReason = 'Night Attendance (Late, No Reason)';
                                              updatedStudents[studentIndex] = st;
                                              studentsToUpsert.push(st);
                                              logHistory(s.id, 'tally', 3, 'Night Attendance (Late, No Reason)');
                                            }
                                          }
                                        }
                                        
                                        updatedRecords[s.id] = { ...rec, tallied: true };
                                        hasChanges = true;
                                      });
                                      
                                      if (studentsToUpsert.length > 0) {
                                        setStudents(updatedStudents);
                                        fetch('/api/students/bulk-upsert', {
                                          method: 'POST',
                                          headers: { 'Content-Type': 'application/json' },
                                          body: JSON.stringify({ students: studentsToUpsert })
                                        }).catch(err => console.error("Error bulk upserting attendance tallies:", err));
                                      }
                                      
                                      if (hasChanges) {
                                        setAttendanceRecords(updatedRecords);
                                      }
                                    }}
                                    className="w-full py-3 rounded-xl text-xs font-bold text-white bg-[#1A365D] hover:bg-[#2A4365] transition-colors shadow-md"
                                  >
                                    Submit Attendance
                                  </button>
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center gap-3 animate-fade-in w-full h-full">
                        <div className="flex items-center justify-between w-full mb-2 shrink-0">
                          <div className="flex items-center gap-2">
                            <span className="text-xl">📊</span>
                            <h3 className="text-sm font-extrabold text-[#1A365D]">Student Summary</h3>
                          </div>
                          <div className="flex gap-2">
                            <button
                              onClick={() => {
                                setIrAssignedStudents(assignedStudents);
                                setDateModalNextAction('IR');
                                setShowIRDateModal(true);
                              }}
                              className="px-3 py-1.5 bg-blue-500 hover:bg-blue-600 text-white text-[10px] uppercase font-bold rounded-lg shadow-sm transition-colors flex items-center gap-1"
                            >
                              <span className="text-sm">📄</span> (IR) Individual Report
                            </button>
                            <button
                              onClick={() => {
                                const csvLines = ["Class,Name,Star,Tally,Total,Grade,N&O Tally,N&O Total,N&O Grade,Diary Tally,Fine,Attitude Total,Attitude Grade"];
                                assignedStudents.forEach(s => {
                                  const total1 = ((Number(s.star) || 0) * 2) - (Number(s.tally) || 0);
                                  const grade = calculateGrade(total1);
                                  const attitudeTotal = ((Number(s.diaryTally) || 0) * -0.5) + (getFineCount(s) * -1.5) + (Number(s.sheetTally) || 0);
                                  const noIncidents = getNOIncidents(s);
                                  const noTotal = -noIncidents;
                                  const noGrade = calculateNOGrade(noTotal);
                                  csvLines.push(`${s.class},"${s.name.replace(/"/g, '""')}",${s.star || 0},${s.tally || 0},${total1},${grade},${s.neatAndOrderTally || 0},${noTotal},${noGrade},${s.diaryTally || 0},${s.fine || 0},${attitudeTotal},${calculateAttitudeGrade(attitudeTotal)}`);
                                });
                                const csvContent = csvLines.join('\n');
                                const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                                const url = URL.createObjectURL(blob);
                                const link = document.createElement("a");
                                link.setAttribute("href", url);
                                link.setAttribute("download", `${mentor.name}_summary.csv`);
                                document.body.appendChild(link);
                                link.click();
                                document.body.removeChild(link);
                              }}
                              className="px-3 py-1.5 bg-emerald-500 hover:bg-emerald-600 text-white text-[10px] uppercase font-bold rounded-lg shadow-sm transition-colors flex items-center gap-1"
                            >
                              <span className="text-sm">📥</span> Download Excel
                            </button>
                            <button
                              onClick={() => {
                                if (confirm("Are you sure you want to clear Star and Tally data for these students? Names and classes will NOT be deleted.")) {
                                  const idsToClear = assignedStudents.map(s => s.id);
                                  setStudents(prev => prev.map(s => {
                                    if (idsToClear.includes(s.id)) {
                                      return { ...s, star: 0, tally: 0, starReason: '', tallyReason: '', diaryStar: 0, diaryTally: 0, neatAndOrderTally: 0, neatAndOrderReason: '', neatAndOrderIncidents: 0 };
                                    }
                                    return s;
                                  }));
                                  
                                  setDiaryRecords(prev => {
                                    const next = { ...prev };
                                    idsToClear.forEach(id => delete next[id]);
                                    return next;
                                  });
                                  
                                  setAttendanceRecords(prev => {
                                    const next = { ...prev };
                                    idsToClear.forEach(id => delete next[id]);
                                    return next;
                                  });

                                  fetch('/api/students/bulk-upsert', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/json' },
                                    body: JSON.stringify({ 
                                      students: assignedStudents.map(s => ({ ...s, star: 0, tally: 0, starReason: '', tallyReason: '', diaryStar: 0, diaryTally: 0, neatAndOrderTally: 0, neatAndOrderReason: '', neatAndOrderIncidents: 0 }))
                                    })
                                  }).catch(err => console.error("Error clearing data:", err));
                                }
                              }}
                              className="px-3 py-1.5 bg-rose-500 hover:bg-rose-600 text-white text-[10px] uppercase font-bold rounded-lg shadow-sm transition-colors flex items-center gap-1"
                            >
                              <span className="text-sm">🗑️</span> Clear Data
                            </button>
                          </div>
                        </div>
                        
                        <div className="w-full flex-1 flex flex-col gap-2 mx-auto overflow-hidden bg-white border border-slate-200 rounded-xl shadow-sm">
                          <div className="overflow-x-auto h-full">
                            <table className="w-full text-left text-xs">
                              <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider sticky top-0 border-b border-slate-200 z-10 shadow-sm">
                                <tr>
                                  <th className="p-3 font-extrabold text-center">Class</th>
                                  <th className="p-3 font-extrabold text-center">Name</th>
                                  <th className="p-3 font-extrabold text-center text-amber-700">Stars</th>
                                  <th className="p-3 font-extrabold text-center text-sky-700">Tallies</th>
                                  <th className="p-3 font-extrabold text-center text-[#1A365D]">Total</th>
                                  <th className="p-3 font-extrabold text-center text-purple-700">Grade</th>
                                  <th className="p-3 font-extrabold text-center text-orange-600">N&O Tally</th>
                                  <th className="p-3 font-extrabold text-center text-[#1A365D]">Total</th>
                                  <th className="p-3 font-extrabold text-center text-orange-700">N&O Grade</th>
                                  <th className="p-3 font-extrabold text-center text-sky-600">Diary Tallies</th>
                                  <th className="p-3 font-extrabold text-center text-cyan-600">Sheets</th>
                                  <th className="p-3 font-extrabold text-center text-rose-600">Fine</th>
                                  <th className="p-3 font-extrabold text-center text-[#1A365D]">Total</th>
                                  <th className="p-3 font-extrabold text-center text-indigo-700">Attitude Grade</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {assignedStudents.length > 0 ? assignedStudents.map(student => {
                                  const total1 = ((Number(student.star) || 0) * 2) - (Number(student.tally) || 0);
                                  const attitudeTotal = ((Number(student.diaryTally) || 0) * -0.5) + (getFineCount(student) * -1.5) + (Number(student.sheetTally) || 0);
                                  return (
                                  <tr key={student.id} className="hover:bg-slate-50 transition-colors">
                                    <td className="p-3 text-center font-bold text-[#1A365D] uppercase">{student.class}</td>
                                    <td className="p-3 font-bold text-slate-700 flex items-center justify-center gap-1.5 h-[45px]">
                                      {student.ineligible && <div className="w-2 h-2 rounded-full bg-rose-600 animate-pulse" title={`Ineligible: ${student.ineligibleReason || 'No reason'}`} />}
                                      {student.name}
                                    </td>
                                    <td className="p-3 text-center font-bold text-amber-600">{student.star || 0}</td>
                                    <td className="p-3 text-center font-bold text-sky-600">{student.tally || 0}</td>
                                    <td className="p-3 text-center font-extrabold text-[#1A365D]">{total1}</td>
                                    <td className="p-3 text-center font-bold text-purple-600">{calculateGrade(total1)}</td>
                                    <td className="p-3 text-center font-bold text-orange-500">{student.neatAndOrderTally || 0}</td>
                                    <td className="p-3 text-center font-extrabold text-[#1A365D]">{-getNOIncidents(student)}</td>
                                    <td className="p-3 text-center font-bold text-orange-600">{calculateNOGrade(-getNOIncidents(student))}</td>
                                    <td className="p-3 text-center font-bold text-sky-500">{student.diaryTally || 0}</td>
                                    <td className="p-3 text-center font-bold text-cyan-500">{student.sheetTally || 0}</td>
                                    <td className="p-3 text-center font-bold text-rose-600">{getFineCount(student)}</td>
                                    <td className="p-3 text-center font-extrabold text-[#1A365D]">{attitudeTotal}</td>
                                    <td className="p-3 text-center font-bold text-indigo-600">
                                      {calculateAttitudeGrade(attitudeTotal)}
                                    </td>
                                  </tr>
                                ); }) : (
                                  <tr>
                                    <td colSpan="14" className="p-6 text-center text-slate-400 font-medium">No students assigned.</td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}
          </div>
        ) : activeTab === 'performance' ? (
          /* PERFORMANCE ANALYTICS VIEW */
          <div className="flex-1 flex flex-col overflow-hidden bg-[#F8F9FA]">
            {!performanceView ? (
              <div className="w-full h-full overflow-y-auto p-4 flex flex-col">
                <div className="w-full max-w-sm mx-auto animate-fade-in flex flex-col gap-4 pb-12 mt-auto mb-auto">
                  <div className="text-center mb-2 mt-4">
                    <h2 className="text-[#1A365D] font-extrabold text-xl tracking-tight">Select Action</h2>
                    <p className="text-slate-500 text-[11px] mt-1">Choose a module to continue</p>
                  </div>
                  
                  <div className="grid grid-cols-2 gap-3 w-full">
                  {/* Neat and Order */}
                  <button 
                    onClick={() => { setPerformanceView('neat'); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="group relative flex flex-col items-center justify-center p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-blue-300 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-3 rounded-2xl bg-blue-50 text-blue-600 mb-3 group-hover:scale-110 group-hover:bg-blue-100 transition-all duration-300">
                      <CheckCircle2 className="w-7 h-7" />
                    </div>
                    <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wide">Neat & Order</h3>
                  </button>

                  {/* Room */}
                  <button 
                    onClick={() => { setPerformanceView('room'); setSelectedHostel(null); setSelectedRoom(null); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="group relative flex flex-col items-center justify-center p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-emerald-300 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-3 rounded-2xl bg-emerald-50 text-emerald-600 mb-3 group-hover:scale-110 group-hover:bg-emerald-100 transition-all duration-300">
                      <School className="w-7 h-7" />
                    </div>
                    <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wide">Room</h3>
                  </button>

                  {/* Spot Fine */}
                  <button 
                    onClick={() => { setPerformanceView('spot'); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="group relative flex flex-col items-center justify-center p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-amber-300 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-3 rounded-2xl bg-amber-50 text-amber-600 mb-3 group-hover:scale-110 group-hover:bg-amber-100 transition-all duration-300">
                      <AlertCircle className="w-7 h-7" />
                    </div>
                    <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wide">Spot Fine</h3>
                  </button>

                  {/* Program Star */}
                  <button 
                    onClick={() => { setPerformanceView('program'); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="group relative flex flex-col items-center justify-center p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-purple-300 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-3 rounded-2xl bg-purple-50 text-purple-600 mb-3 group-hover:scale-110 group-hover:bg-purple-100 transition-all duration-300">
                      <Star className="w-7 h-7" />
                    </div>
                    <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wide">Prog. Star</h3>
                  </button>

                  {/* Ineligible */}
                  <button 
                    onClick={() => { setPerformanceView('ineligible'); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="group relative flex flex-col items-center justify-center p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-rose-300 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-3 rounded-2xl bg-rose-50 text-rose-600 mb-3 group-hover:scale-110 group-hover:bg-rose-100 transition-all duration-300">
                      <X className="w-7 h-7" />
                    </div>
                    <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wide">Ineligible</h3>
                  </button>

                  {/* Sheets */}
                  <button 
                    onClick={() => { setPerformanceView('sheets'); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="group relative flex flex-col items-center justify-center p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-cyan-300 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-3 rounded-2xl bg-cyan-50 text-cyan-600 mb-3 group-hover:scale-110 group-hover:bg-cyan-100 transition-all duration-300">
                      <Layers className="w-7 h-7" />
                    </div>
                    <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wide">Sheets</h3>
                  </button>

                  {/* Morning Bliss */}
                  <button 
                    onClick={() => { setPerformanceView('morning_bliss'); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="group relative flex flex-col items-center justify-center p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-pink-300 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-3 rounded-2xl bg-pink-50 text-pink-600 mb-3 group-hover:scale-110 group-hover:bg-pink-100 transition-all duration-300">
                      <Sun className="w-7 h-7" />
                    </div>
                    <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wide">Morning Bliss</h3>
                  </button>

                  {/* Morning Bliss Results */}
                  <button 
                    onClick={() => { setPerformanceView('morning_bliss_results'); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="group relative flex flex-col items-center justify-center p-5 bg-white border border-slate-200 rounded-2xl shadow-sm hover:shadow-md hover:border-violet-300 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-3 rounded-2xl bg-violet-50 text-violet-600 mb-3 group-hover:scale-110 group-hover:bg-violet-100 transition-all duration-300">
                      <List className="w-7 h-7" />
                    </div>
                    <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wide">MB Results</h3>
                  </button>

                  {/* Summary */}
                  <button 
                    onClick={() => { setPerformanceView('summary'); setPerformanceSelectedClass(null); setPerformanceSelectedStudents([]); }}
                    className="col-span-2 group relative flex flex-row items-center justify-center p-4 bg-slate-800 border border-slate-800 rounded-2xl shadow-md hover:shadow-lg hover:bg-slate-900 transition-all duration-200 active:scale-95 text-center"
                  >
                    <div className="p-2 rounded-xl bg-slate-700/50 text-white mr-3 group-hover:scale-110 transition-all duration-300">
                      <FileText className="w-5 h-5" />
                    </div>
                    <h3 className="font-extrabold text-white text-sm uppercase tracking-wider">View Summary</h3>
                  </button>
                </div>
              </div>
             </div>
            ) : (
              <div className="flex flex-col w-full h-full">
                <div className="p-4 bg-white border-b border-slate-200 shadow-sm shrink-0 flex items-center gap-3">
                  <button 
                    onClick={() => {
                      if (performanceSelectedStudents.length > 0) {
                        setPerformanceSelectedStudents([]);
                      } else if (performanceView === 'room' && selectedRoom) {
                        setSelectedRoom(null);
                      } else if (performanceView === 'room' && selectedHostel) {
                        setSelectedHostel(null);
                      } else if (performanceSelectedClass) {
                        setPerformanceSelectedClass(null);
                      } else if (performanceView && performanceView.startsWith('sheets_')) {
                        setPerformanceView('sheets');
                      } else {
                        setPerformanceView(null);
                        setSelectedHostel(null);
                        setSelectedRoom(null);
                      }
                    }}
                    className="p-2 bg-slate-50 hover:bg-slate-100 rounded-xl transition-colors text-[#1A365D]"
                  >
                    <ChevronLeft className="w-5 h-5" />
                  </button>
                  <h2 className="text-[#1A365D] font-extrabold tracking-wider uppercase text-sm">
                    {performanceView === 'neat' ? 'Neat and Order' : performanceView === 'room' ? (selectedRoom ? `Room - ${selectedHostel} (${selectedRoom})` : selectedHostel ? `Room - ${selectedHostel}` : 'Hostel Rooms') : performanceView === 'spot' ? 'Spot Fine' : performanceView === 'program' ? 'Program Star' : performanceView === 'ineligible' ? 'Ineligible' : performanceView === 'sheets' ? 'Sheets' : performanceView === 'sheets_black' ? 'Black Sheet' : performanceView === 'sheets_yellow' ? 'Yellow Sheet' : performanceView === 'sheets_apology' ? 'Apology Sheet' : performanceView === 'morning_bliss' ? 'Morning Bliss' : performanceView === 'morning_bliss_results' ? 'Morning Bliss Results' : performanceView === 'summary' ? 'Summary' : ''}
                    {performanceSelectedClass ? ` - Class ${performanceSelectedClass}` : ''}
                  </h2>
                </div>

                {performanceView === 'morning_bliss' ? (
                  <div className="flex-1 overflow-y-auto p-4 bg-slate-50 relative pb-24">
                    <form onSubmit={handleMorningBlissSubmit} className="max-w-md mx-auto space-y-4">
                      {/* Class Selection */}
                      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                        <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Class</label>
                        <select 
                          value={morningBlissClass}
                          onChange={(e) => { setMorningBlissClass(e.target.value); setMorningBlissNameSearch(''); }}
                          className="w-full bg-slate-50 border border-slate-200 text-slate-800 rounded-lg p-3 font-semibold outline-none focus:ring-2 focus:ring-pink-500 focus:border-transparent transition-all"
                          required
                        >
                          <option value="">Select Class</option>
                          {CLASSES.map(c => <option key={c} value={c}>{c.toUpperCase()}</option>)}
                        </select>
                      </div>

                      {/* Name Search */}
                      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm relative">
                        <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Student Name</label>
                        <input
                          type="text"
                          placeholder="Enter name..."
                          value={morningBlissNameSearch}
                          onChange={(e) => { setMorningBlissNameSearch(e.target.value); setShowMorningBlissDropdown(true); }}
                          onFocus={() => setShowMorningBlissDropdown(true)}
                          onBlur={() => setTimeout(() => setShowMorningBlissDropdown(false), 200)}
                          className="w-full bg-slate-50 border border-slate-200 text-slate-800 rounded-lg p-3 font-semibold outline-none focus:ring-2 focus:ring-pink-500 focus:border-transparent transition-all"
                          required
                        />
                        {showMorningBlissDropdown && morningBlissClass && morningBlissNameSearch.length > 0 && (
                          <div className="absolute z-10 w-full mt-1 bg-white border border-slate-200 rounded-lg shadow-xl max-h-48 overflow-y-auto">
                            {students.filter(s => s.class === morningBlissClass && s.name.toLowerCase().includes(morningBlissNameSearch.toLowerCase())).map(s => (
                              <div
                                key={s.id}
                                className="p-3 hover:bg-pink-50 cursor-pointer font-medium text-slate-700 border-b border-slate-100 last:border-0"
                                onClick={() => { setMorningBlissNameSearch(s.name); setShowMorningBlissDropdown(false); }}
                              >
                                {s.name}
                              </div>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* Topic & Mark & EV */}
                      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-4">
                        <div>
                          <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Topic</label>
                          <input
                            type="text"
                            placeholder="Enter Topic..."
                            value={morningBlissTopic}
                            onChange={(e) => setMorningBlissTopic(e.target.value)}
                            className="w-full bg-slate-50 border border-slate-200 text-slate-800 rounded-lg p-3 font-semibold outline-none focus:ring-2 focus:ring-pink-500 focus:border-transparent transition-all"
                            required
                          />
                        </div>
                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Mark (Max 8)</label>
                            <input
                              type="number"
                              min="0"
                              max="8"
                              step="0.1"
                              placeholder="0 - 8"
                              value={morningBlissMark}
                              onChange={(e) => setMorningBlissMark(e.target.value)}
                              className="w-full bg-slate-50 border border-slate-200 text-slate-800 rounded-lg p-3 font-semibold outline-none focus:ring-2 focus:ring-pink-500 focus:border-transparent transition-all"
                              required
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">EV (Teacher)</label>
                            <input
                              type="text"
                              placeholder="EV Name..."
                              value={morningBlissEv}
                              onChange={(e) => setMorningBlissEv(e.target.value)}
                              className="w-full bg-slate-50 border border-slate-200 text-slate-800 rounded-lg p-3 font-semibold outline-none focus:ring-2 focus:ring-pink-500 focus:border-transparent transition-all"
                              required
                            />
                          </div>
                        </div>
                      </div>

                      {/* Stopwatch */}
                      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm text-center">
                        <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Stopwatch</label>
                        <div className="text-4xl font-black text-slate-800 tracking-widest font-mono mb-4">
                          {formatMBTime(morningBlissDuration)}
                        </div>
                        <div className="flex gap-2 justify-center">
                          <button
                            type="button"
                            onClick={() => setIsStopwatchRunning(!isStopwatchRunning)}
                            className={`px-4 py-2 rounded-lg font-bold text-white transition-colors ${isStopwatchRunning ? 'bg-red-500 hover:bg-red-600' : 'bg-emerald-500 hover:bg-emerald-600'}`}
                          >
                            {isStopwatchRunning ? 'Stop' : 'Start'}
                          </button>
                          <button
                            type="button"
                            onClick={() => { setIsStopwatchRunning(false); setMorningBlissDuration(0); }}
                            className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg font-bold transition-colors"
                          >
                            Reset
                          </button>
                        </div>
                      </div>

                      {/* Submit */}
                      <button
                        type="submit"
                        className="w-full bg-pink-600 hover:bg-pink-700 text-white font-extrabold p-4 rounded-xl shadow-lg hover:shadow-xl transition-all duration-300 active:scale-95 flex items-center justify-center gap-2 uppercase tracking-widest"
                      >
                        <CheckCircle2 className="w-5 h-5" />
                        Done and Return
                      </button>
                    </form>
                  </div>
                ) : performanceView === 'morning_bliss_results' ? (
                  <div className="flex-1 overflow-hidden flex flex-col">
                    <div className="p-4 bg-white border-b border-slate-200 shrink-0 flex flex-wrap justify-between items-center gap-3">
                      <div className="flex flex-col">
                        <span className="font-bold text-slate-800">Today's Morning Bliss Results</span>
                        <span className="text-xs font-medium text-slate-500">Summary Date: {new Date().toISOString().split('T')[0]}</span>
                      </div>
                      
                      <div className="flex flex-wrap items-center gap-3">
                        {/* Excel Export with Date Range Filter */}
                        <div className="flex items-center gap-2 bg-slate-100 p-2 rounded-lg text-xs font-semibold">
                          <span className="text-slate-600 font-bold">Range:</span>
                          <label className="flex items-center gap-1 text-slate-700">
                            From
                            <input 
                              type="date" 
                              value={mbFromDate} 
                              onChange={(e) => setMbFromDate(e.target.value)} 
                              className="p-1 border rounded bg-white text-slate-800 outline-none focus:ring-1 focus:ring-emerald-500"
                            />
                          </label>
                          <label className="flex items-center gap-1 text-slate-700">
                            To
                            <input 
                              type="date" 
                              value={mbToDate} 
                              onChange={(e) => setMbToDate(e.target.value)} 
                              className="p-1 border rounded bg-white text-slate-800 outline-none focus:ring-1 focus:ring-emerald-500"
                            />
                          </label>
                          <button 
                            onClick={handleMorningBlissExportExcelRange}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-md font-bold transition-colors shadow-sm flex items-center gap-1"
                            title="Export Excel for selected Summary Date range"
                          >
                            <Download className="w-4 h-4" /> Excel
                          </button>
                        </div>

                        <button 
                          onClick={() => {
                            if (window.confirm("Are you sure you want to clear all Morning Bliss data for today?")) {
                              const todayDateStr = new Date().toISOString().split('T')[0];
                              const studentsToClear = students.filter(s => 
                                s.summaryDate === todayDateStr || 
                                s.morningBlissMark != null || 
                                s.morningBlissScript != null || 
                                s.morningBlissTopic
                              );
                              if (studentsToClear.length === 0) return;
                              
                              const updatedStudents = students.map(s => {
                                if (
                                  s.summaryDate === todayDateStr || 
                                  s.morningBlissMark != null || 
                                  s.morningBlissScript != null || 
                                  s.morningBlissTopic
                                ) {
                                  return { 
                                    ...s, 
                                    morningBlissMark: null, 
                                    morningBlissScript: null, 
                                    morningBlissTopic: '', 
                                    morningBlissEv: '', 
                                    morningBlissDuration: '', 
                                    morningBlissStar: 0,
                                    summaryId: null,
                                    summaryDate: null
                                  };
                                }
                                return s;
                              });
                              
                              setStudents(updatedStudents);
                              
                              fetch('/api/students/bulk-upsert', {
                                method: 'POST',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify({ 
                                  students: updatedStudents.filter(s => studentsToClear.map(stu => stu.id).includes(s.id))
                                })
                              }).catch(err => console.error("Error bulk clearing morning bliss:", err));
                            }
                          }}
                          className="bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-lg font-bold text-sm transition-colors shadow-sm"
                        >
                          CLEAR
                        </button>
                        <button 
                          onClick={handleMorningBlissDone}
                          className="bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-lg font-bold text-sm transition-colors shadow-sm flex items-center gap-1"
                        >
                          <FileText className="w-4 h-4" /> PDF Report
                        </button>
                      </div>
                    </div>
                    <div className="flex-1 overflow-y-auto p-4 bg-slate-50">
                      <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="bg-slate-100 border-b border-slate-200">
                                <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Class</th>
                                <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Name</th>
                                <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Mark</th>
                                <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Script Mark</th>
                                <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Total</th>
                                <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Star</th>
                                <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">EV</th>
                              </tr>
                            </thead>
                            <tbody>
                              {(() => {
                                const todayDateStr = new Date().toISOString().split('T')[0];
                                const rows = CLASSES.flatMap(c => {
                                  const classStudents = students.filter(s => s.class === c && s.morningBlissMark != null && s.morningBlissMark !== '' && (s.summaryDate === todayDateStr || !s.summaryDate));
                                  if (classStudents.length === 0) return [];

                                  return classStudents.map(s => {
                                    const total = (Number(s.morningBlissMark) || 0) + (Number(s.morningBlissScript) || 0);
                                    return (
                                    <tr key={s.id} className="border-b border-slate-100 hover:bg-slate-50">
                                      <td className="p-3 font-bold text-slate-800">{s.class.toUpperCase()}</td>
                                      <td className="p-3 font-medium text-slate-700">{s.name}</td>
                                      <td className="p-3 font-bold text-indigo-600">{s.morningBlissMark != null ? s.morningBlissMark : '-'}</td>
                                      <td className="p-3">
                                        <input 
                                          type="number"
                                          min="0"
                                          max="2"
                                          step="0.1"
                                          value={s.morningBlissScript != null ? s.morningBlissScript : ''}
                                          onChange={(e) => {
                                            const val = e.target.value === '' ? null : Number(e.target.value);
                                            
                                            const newTotal = (Number(s.morningBlissMark) || 0) + (val || 0);
                                            let newStars = 0;
                                            if (newTotal === 10) newStars = 3;
                                            else if (newTotal >= 9.5) newStars = 2;
                                            else if (newTotal >= 9) newStars = 1;

                                            const starDiff = newStars - (s.morningBlissStar || 0);
                                            
                                            const updated = {
                                              ...s,
                                              morningBlissScript: val,
                                              morningBlissStar: newStars,
                                              star: (s.star || 0) + starDiff
                                            };
                                            
                                            setStudents(prev => prev.map(student => student.id === s.id ? updated : student));
                                            
                                            fetch('/api/students/bulk-upsert', {
                                                method: 'POST',
                                                headers: { 'Content-Type': 'application/json' },
                                                body: JSON.stringify({ students: [updated] })
                                            }).catch(err => console.error("Error bulk upserting:", err));
                                          }}
                                          className="w-16 p-1 border rounded bg-white text-center focus:ring-2 focus:ring-violet-500"
                                        />
                                      </td>
                                      <td className="p-3 font-bold text-emerald-600">{total > 0 ? total : '-'}</td>
                                      <td className="p-3 font-bold text-amber-500">{s.morningBlissStar > 0 ? Array(s.morningBlissStar).fill('⭐').join('') : '-'}</td>
                                      <td className="p-3 text-slate-600 text-sm">{s.morningBlissEv || '-'}</td>
                                    </tr>
                                  )});
                                });

                                return rows.length === 0 ? (
                                  <tr>
                                    <td colSpan="7" className="p-8 text-center text-slate-400 font-medium">
                                      No Morning Bliss evaluations recorded for today yet. Use the "Morning Bliss" form to submit evaluations for today.
                                    </td>
                                  </tr>
                                ) : (
                                  rows
                                );
                              })()}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                ) : performanceView === 'summary' ? (
                  <div className="flex-1 overflow-hidden flex flex-col">
                    <div className="p-4 flex items-center justify-between bg-white border-b border-slate-200 shrink-0">
                       <span className="font-bold text-slate-700">All Tallies Summary</span>
                         <button
                         onClick={() => {
                           const ws = XLSX.utils.json_to_sheet(students.filter(s => s.tally > 0 || s.neatAndOrderTally > 0 || s.fine > 0).map(s => ({
                             Class: s.class.toUpperCase(),
                             Name: s.name,
                             Tallies: s.tally || 0,
                             'N&O Tallies': s.neatAndOrderTally || 0,
                             Fine: s.fine || 0,
                             Reason: s.tallyReason || s.neatAndOrderReason || s.fineReason
                           })));
                           const wb = XLSX.utils.book_new();
                           XLSX.utils.book_append_sheet(wb, ws, "Tallies");
                           XLSX.writeFile(wb, "Tallies_Summary.xlsx");
                         }}
                         className="px-3 py-1.5 bg-emerald-600 text-white text-xs font-bold rounded-lg flex items-center gap-2 shadow-sm hover:bg-emerald-700 active:scale-95 transition-all"
                       >
                         <Download className="w-4 h-4" /> Export Excel
                       </button>
                    </div>
                    <div className="flex-1 overflow-y-auto p-4">
                      <table className="w-full text-left text-xs bg-white rounded shadow-sm border border-slate-200">
                        <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase tracking-wider">
                          <tr>
                            <th className="p-3 font-extrabold">Class</th>
                            <th className="p-3 font-extrabold">Name</th>
                            <th className="p-3 font-extrabold text-center">Tally</th>
                            <th className="p-3 font-extrabold text-center">N&O Tally</th>
                            <th className="p-3 font-extrabold text-center">Fine</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {students.filter(s => s.tally > 0 || s.neatAndOrderTally > 0 || s.fine > 0).map(s => (
                            <tr key={s.id} className="hover:bg-slate-50 transition-colors">
                              <td className="p-3 uppercase font-bold text-[#1A365D]">{s.class}</td>
                              <td className="p-3 font-semibold text-slate-700">{s.name}</td>
                              <td className="p-3 text-center text-rose-500 font-extrabold">{s.tally || 0}</td>
                              <td className="p-3 text-center text-orange-500 font-extrabold">{s.neatAndOrderTally || 0}</td>
                              <td className="p-3 text-center text-rose-600 font-extrabold">{s.fine || 0}</td>
                            </tr>
                          ))}
                          {students.filter(s => s.tally > 0 || s.neatAndOrderTally > 0 || s.fine > 0).length === 0 && (
                            <tr>
                              <td colSpan="5" className="p-6 text-center text-slate-400 font-medium">No tallies recorded.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : performanceView === 'ineligible' ? (
                  <div className="flex-1 overflow-hidden flex flex-col">
                    <div className="p-4 flex items-center justify-between bg-white border-b border-slate-200 shrink-0">
                       <span className="font-bold text-slate-700">Ineligible Students</span>
                       <button
                         onClick={() => setShowAddIneligibleModal(true)}
                         className="px-4 py-2 bg-rose-600 text-white text-xs font-bold rounded-xl flex items-center gap-2 shadow-sm hover:bg-rose-700 active:scale-95 transition-all"
                       >
                         <Plus className="w-4 h-4" /> Add Ineligible
                       </button>
                    </div>
                    <div className="flex-1 overflow-y-auto p-4">
                      <table className="w-full text-left text-xs bg-white rounded shadow-sm border border-slate-200">
                        <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase tracking-wider">
                          <tr>
                            <th className="p-3 font-extrabold w-24">Class</th>
                            <th className="p-3 font-extrabold">Name</th>
                            <th className="p-3 font-extrabold">Reason</th>
                            <th className="p-3 font-extrabold text-right">Action</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {students.filter(s => s.ineligible).map(s => (
                            <tr key={s.id} className="hover:bg-slate-50 transition-colors">
                              <td className="p-3 uppercase font-bold text-[#1A365D]">{s.class}</td>
                              <td className="p-3 font-extrabold text-rose-600 uppercase tracking-widest">{s.name}</td>
                              <td className="p-3 font-medium text-slate-700">{s.ineligibleReason || '-'}</td>
                              <td className="p-3 text-right">
                                <button 
                                  onClick={() => handleRemoveIneligible(s.id)}
                                  className="text-[10px] font-bold text-slate-400 hover:text-slate-600 transition-colors uppercase tracking-wider"
                                >
                                  Remove
                                </button>
                              </td>
                            </tr>
                          ))}
                          {students.filter(s => s.ineligible).length === 0 && (
                            <tr>
                              <td colSpan="4" className="p-6 text-center text-slate-400 font-medium">No ineligible students.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : performanceView === 'sheets' ? (
                  <div className="w-full max-w-sm p-4 mx-auto animate-fade-in flex flex-col gap-4 mt-8">
                    <div className="text-center mb-4">
                      <h2 className="text-[#1A365D] font-extrabold text-xl tracking-tight">Select Sheet Type</h2>
                      <p className="text-slate-500 text-[11px] mt-1">Choose the type of sheet to submit</p>
                    </div>
                    <div className="flex flex-col gap-3">
                      <button 
                        onClick={() => setPerformanceView('sheets_black')}
                        className="p-4 bg-slate-900 border border-slate-900 rounded-2xl shadow-sm hover:shadow-md hover:bg-slate-800 transition-all text-white text-center font-extrabold text-sm active:scale-[0.95]"
                      >
                        Black Sheet
                      </button>
                      <button 
                        onClick={() => setPerformanceView('sheets_yellow')}
                        className="p-4 bg-amber-400 border border-amber-400 rounded-2xl shadow-sm hover:shadow-md hover:bg-amber-300 transition-all text-slate-900 text-center font-extrabold text-sm active:scale-[0.95]"
                      >
                        Yellow Sheet
                      </button>
                      <button 
                        onClick={() => setPerformanceView('sheets_apology')}
                        className="p-4 bg-blue-100 border border-blue-200 rounded-2xl shadow-sm hover:shadow-md hover:bg-blue-200 transition-all text-blue-900 text-center font-extrabold text-sm active:scale-[0.95]"
                      >
                        Apology
                      </button>
                    </div>
                  </div>
                ) : performanceView?.startsWith('sheets_') ? (
                  <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                    <form onSubmit={handleSheetSubmit} className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col gap-4">
                        <div>
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Select Class</label>
                            <select 
                                value={sheetClass} 
                                onChange={(e) => setSheetClass(e.target.value)}
                                className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50 appearance-none"
                                required
                            >
                                <option value="" disabled>Choose a class...</option>
                                {visibleClasses.map(cls => <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>)}
                            </select>
                        </div>
                        <div className="relative">
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Student Name</label>
                            <div className="relative">
                                <Search className="w-5 h-5 text-slate-400 absolute left-4 top-1/2 -translate-y-1/2" />
                                <input 
                                    type="text" 
                                    value={sheetNameSearch} 
                                    onChange={(e) => {
                                        setSheetNameSearch(e.target.value);
                                        setShowSheetDropdown(true);
                                    }}
                                    onFocus={() => setShowSheetDropdown(true)}
                                    placeholder="Type student name..."
                                    className="w-full p-4 pl-12 text-lg border border-slate-200 rounded-xl font-bold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                                    required
                                />
                            </div>
                            {sheetNameSearch && sheetClass && showSheetDropdown && (
                                <div className="mt-2 bg-white border border-slate-200 rounded-xl shadow-sm max-h-60 overflow-y-auto z-10 absolute left-0 right-0">
                                    {students.filter(s => s.class === sheetClass && s.name.toLowerCase().startsWith(sheetNameSearch.toLowerCase())).map(s => (
                                        <div 
                                            key={s.id} 
                                            className="p-4 hover:bg-slate-50 cursor-pointer font-bold text-slate-700 border-b border-slate-100 last:border-0 flex items-center justify-between transition-colors group"
                                            onClick={() => {
                                                setSheetNameSearch(s.name);
                                                setShowSheetDropdown(false);
                                            }}
                                        >
                                            {s.name}
                                            <CheckCircle2 className="w-5 h-5 text-emerald-500 opacity-0 group-hover:opacity-100 transition-opacity" />
                                        </div>
                                    ))}
                                    {students.filter(s => s.class === sheetClass && s.name.toLowerCase().startsWith(sheetNameSearch.toLowerCase())).length === 0 && (
                                        <div className="p-4 text-center text-slate-400 font-semibold text-sm">
                                            No matching student found in Class {sheetClass.toUpperCase()}
                                        </div>
                                    )}
                                </div>
                            )}
                        </div>
                        <div>
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Reason</label>
                            <input 
                                type="text" 
                                value={sheetReason} 
                                onChange={(e) => setSheetReason(e.target.value)}
                                placeholder="E.g. Disruption, late coming..."
                                className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                                required
                            />
                        </div>
                        <button type="submit" className="w-full py-4 mt-2 bg-blue-600 hover:bg-blue-700 text-white font-extrabold rounded-xl transition-all shadow-md active:scale-[0.98] flex justify-center items-center gap-2">
                            <Send className="w-5 h-5" /> Submit Sheet
                        </button>
                    </form>
                  </div>
                ) : performanceView === 'spot' ? (
                  <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                    <form onSubmit={handleSpotFineSubmit} className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col gap-4">
                        <div>
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Select Class</label>
                            <select 
                                value={spotClass} 
                                onChange={(e) => setSpotClass(e.target.value)}
                                className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50 appearance-none"
                                required
                            >
                                <option value="" disabled>Choose a class...</option>
                                {visibleClasses.map(cls => <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>)}
                            </select>
                        </div>
                        <div className="relative">
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Student Name</label>
                            <div className="relative">
                                <Search className="w-5 h-5 text-slate-400 absolute left-4 top-1/2 -translate-y-1/2" />
                                <input 
                                    type="text" 
                                    value={spotNameSearch} 
                                    onChange={(e) => {
                                        setSpotNameSearch(e.target.value);
                                        setShowSpotDropdown(true);
                                    }}
                                    onFocus={() => setShowSpotDropdown(true)}
                                    placeholder="Type student name..."
                                    className="w-full p-4 pl-12 text-lg border border-slate-200 rounded-xl font-bold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                                    required
                                />
                            </div>
                            {spotNameSearch && spotClass && showSpotDropdown && (
                                <div className="mt-2 bg-white border border-slate-200 rounded-xl shadow-sm max-h-60 overflow-y-auto">
                                    {students.filter(s => s.class === spotClass && s.name.toLowerCase().startsWith(spotNameSearch.toLowerCase())).map(s => (
                                        <div 
                                            key={s.id} 
                                            className="p-4 hover:bg-slate-50 cursor-pointer font-bold text-slate-700 border-b border-slate-100 last:border-0 flex items-center justify-between transition-colors group"
                                            onClick={() => {
                                                setSpotNameSearch(s.name);
                                                setShowSpotDropdown(false);
                                            }}
                                        >
                                            {s.name}
                                            <CheckCircle2 className="w-5 h-5 text-emerald-500 opacity-0 group-hover:opacity-100 transition-opacity" />
                                        </div>
                                    ))}
                                    {students.filter(s => s.class === spotClass && s.name.toLowerCase().startsWith(spotNameSearch.toLowerCase())).length === 0 && (
                                        <div className="p-4 text-center text-slate-400 font-semibold text-sm">
                                            No matching student found in Class {spotClass.toUpperCase()}
                                        </div>
                                    )}
                                </div>
                            )}
                        </div>
                        <div className="relative">
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Reason</label>
                            <input 
                                type="text" 
                                value={spotReason} 
                                onChange={(e) => {
                                  setSpotReason(e.target.value);
                                  setShowSpotReasonDropdown(true);
                                }}
                                onFocus={() => setShowSpotReasonDropdown(true)}
                                onBlur={() => setTimeout(() => setShowSpotReasonDropdown(false), 200)}
                                placeholder="E.g. Disruption, late coming..."
                                className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                            />
                            {showSpotReasonDropdown && spotReason && (
                                (() => {
                                    const filtered = PREDEFINED_SPOT_FINES.filter(f => f.reason.toLowerCase().includes(spotReason.toLowerCase()));
                                    if (filtered.length === 0) return null;
                                    return (
                                        <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-slate-200 rounded-xl shadow-lg max-h-48 overflow-y-auto z-50 divide-y divide-slate-100">
                                            {filtered.map((item, idx) => (
                                                <button
                                                    key={idx}
                                                    type="button"
                                                    onMouseDown={() => {
                                                        setSpotReason(item.reason);
                                                        setSpotAmount(item.amount);
                                                        setShowSpotReasonDropdown(false);
                                                    }}
                                                    className="w-full text-left px-4 py-3 text-sm text-slate-700 hover:bg-[#1A365D]/5 hover:text-[#1A365D] transition-all flex justify-between items-center"
                                                >
                                                    <span className="font-semibold">{item.reason}</span>
                                                    <span className="text-xs font-extrabold px-2 py-1 rounded-full bg-rose-100 border border-rose-200 text-rose-800">
                                                        {item.amount}
                                                    </span>
                                                </button>
                                            ))}
                                        </div>
                                    );
                                })()
                            )}
                        </div>
                        <div>
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Fine Amount</label>
                            <input 
                                type="number" 
                                value={spotAmount} 
                                onChange={(e) => setSpotAmount(e.target.value)}
                                placeholder="Enter fine amount"
                                className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                                required
                            />
                        </div>
                        <button type="submit" className="w-full py-4 mt-2 bg-rose-500 hover:bg-rose-600 text-white font-extrabold rounded-xl transition-all shadow-md active:scale-[0.98] flex justify-center items-center gap-2">
                            <AlertCircle className="w-5 h-5" /> Apply Spot Fine
                        </button>
                    </form>
                  </div>
                ) : performanceView === 'program' ? (
                  <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                    <form onSubmit={handleProgramStarSubmit} className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col gap-4">
                        <div>
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Select Class</label>
                            <select 
                                value={programClass} 
                                onChange={(e) => setProgramClass(e.target.value)}
                                className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50 appearance-none"
                                required
                            >
                                <option value="" disabled>Choose a class...</option>
                                {visibleClasses.map(cls => <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>)}
                            </select>
                        </div>
                        <div className="relative">
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Student Name</label>
                            <div className="relative">
                                <Search className="w-5 h-5 text-slate-400 absolute left-4 top-1/2 -translate-y-1/2" />
                                <input 
                                    type="text" 
                                    value={programNameSearch} 
                                    onChange={(e) => {
                                        setProgramNameSearch(e.target.value);
                                        setShowProgramDropdown(true);
                                    }}
                                    onFocus={() => setShowProgramDropdown(true)}
                                    placeholder="Type student name..."
                                    className="w-full p-4 pl-12 text-lg border border-slate-200 rounded-xl font-bold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                                    required
                                />
                            </div>
                            {programNameSearch && programClass && showProgramDropdown && (
                                <div className="mt-2 bg-white border border-slate-200 rounded-xl shadow-sm max-h-60 overflow-y-auto">
                                    {students.filter(s => s.class === programClass && s.name.toLowerCase().startsWith(programNameSearch.toLowerCase())).map(s => (
                                        <div 
                                            key={s.id} 
                                            className="p-4 hover:bg-slate-50 cursor-pointer font-bold text-slate-700 border-b border-slate-100 last:border-0 flex items-center justify-between transition-colors group"
                                            onClick={() => {
                                                setProgramNameSearch(s.name);
                                                setShowProgramDropdown(false);
                                            }}
                                        >
                                            {s.name}
                                            <CheckCircle2 className="w-5 h-5 text-emerald-500 opacity-0 group-hover:opacity-100 transition-opacity" />
                                        </div>
                                    ))}
                                    {students.filter(s => s.class === programClass && s.name.toLowerCase().startsWith(programNameSearch.toLowerCase())).length === 0 && (
                                        <div className="p-4 text-center text-slate-400 font-semibold text-sm">
                                            No matching student found in Class {programClass.toUpperCase()}
                                        </div>
                                    )}
                                </div>
                            )}
                        </div>
                        <div>
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Programme (Optional)</label>
                            <input 
                                type="text" 
                                value={programReason} 
                                onChange={(e) => setProgramReason(e.target.value)}
                                placeholder="E.g. Outstanding project, leadership..."
                                className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                            />
                        </div>
                        <div>
                            <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Star Count</label>
                            <input 
                                type="number" 
                                value={programAmount} 
                                onChange={(e) => setProgramAmount(e.target.value)}
                                placeholder="Enter star amount"
                                className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                                required
                            />
                        </div>
                        <button type="submit" className="w-full py-4 mt-2 bg-purple-500 hover:bg-purple-600 text-white font-extrabold rounded-xl transition-all shadow-md active:scale-[0.98] flex justify-center items-center gap-2">
                            <Star className="w-5 h-5 fill-white" /> Add Program Star
                        </button>
                    </form>
                  </div>
                ) : performanceView === 'room' ? (
                  !selectedHostel ? (
                    /* STEP 1: Show 4 Hostel Buttons */
                    <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                      <div className="text-center mb-1">
                        <h3 className="text-[#1A365D] font-extrabold text-lg tracking-tight">Select Hostel Block</h3>
                        <p className="text-slate-500 text-xs mt-0.5 font-medium">Choose a hostel to view room list</p>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5 max-w-lg mx-auto w-full">
                        {HOSTEL_BLOCKS.map(h => (
                          <button
                            key={h.id}
                            onClick={() => {
                              setSelectedHostel(h.name);
                              setSelectedRoom(null);
                            }}
                            className="p-5 bg-white border border-slate-200 hover:border-[#1A365D] rounded-2xl shadow-sm hover:shadow-md transition-all active:scale-[0.98] flex items-center justify-between group text-left"
                          >
                            <div className="flex items-center gap-3.5">
                              <div className="w-12 h-12 rounded-2xl bg-slate-100 group-hover:bg-[#1A365D] group-hover:text-white text-2xl flex items-center justify-center transition-colors">
                                {h.icon}
                              </div>
                              <div>
                                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">Hostel {h.id}</span>
                                <span className="font-extrabold text-sm text-[#1A365D] uppercase tracking-wide">{h.name}</span>
                              </div>
                            </div>
                            <ChevronRight className="w-5 h-5 text-slate-400 group-hover:text-[#1A365D] group-hover:translate-x-1 transition-all" />
                          </button>
                        ))}
                      </div>
                    </div>
                  ) : !selectedRoom ? (
                    /* STEP 2: Show Room Buttons for Selected Hostel */
                    <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                      <div className="flex items-center justify-between bg-white p-3.5 rounded-2xl border border-slate-200 shadow-sm shrink-0">
                        <div>
                          <span className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider block">{selectedHostel}</span>
                          <h3 className="text-[#1A365D] font-extrabold text-base">Select Room ({[19, 20, 21, 22, 123, 124, 125, 126, 127, 128, 107, 108, 109, 110, 214, 215, 216, 217, 230, 231, 232, 233, 234, 235].length} Rooms)</h3>
                        </div>
                        <button
                          onClick={() => setSelectedHostel(null)}
                          className="text-xs font-bold text-slate-500 hover:text-slate-800 bg-slate-100 hover:bg-slate-200 px-3 py-1.5 rounded-xl transition-colors"
                        >
                          Change Hostel
                        </button>
                      </div>
                      <div className="grid grid-cols-3 sm:grid-cols-5 gap-3">
                        {[19, 20, 21, 22, 123, 124, 125, 126, 127, 128, 107, 108, 109, 110, 214, 215, 216, 217, 230, 231, 232, 233, 234, 235].map(num => (
                          <button
                            key={num}
                            onClick={() => {
                              const roomStr = `Room ${num}`;
                              setSelectedRoom(roomStr);
                              const allocated = ROOM_STUDENT_MAPPING[String(num)] !== undefined ? 
                                ROOM_STUDENT_MAPPING[String(num)] : 
                                students.filter(s => String(s.room || s.roomNumber || '').trim() === String(num));
                              setPerformanceSelectedStudents(allocated.map(s => s.id));
                            }}
                            className="p-4 bg-white border border-slate-200 hover:border-emerald-500 hover:bg-emerald-50/40 rounded-2xl shadow-sm hover:shadow transition-all active:scale-[0.96] flex flex-col items-center justify-center gap-1.5 group"
                          >
                            <div className="w-9 h-9 rounded-xl bg-slate-100 group-hover:bg-emerald-600 text-slate-600 group-hover:text-white flex items-center justify-center transition-colors">
                              <School className="w-5 h-5" />
                            </div>
                            <span className="font-extrabold text-xs text-slate-800 group-hover:text-emerald-700">Room {num}</span>
                          </button>
                        ))}
                      </div>
                    </div>
                  ) : (
                    /* STEP 3: Student Selection & Room Actions inside Selected Room */
                    (() => {
                      const currentRoomKey = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
                      const allocatedRoomStudents = ROOM_STUDENT_MAPPING[currentRoomKey] !== undefined ? 
                        ROOM_STUDENT_MAPPING[currentRoomKey] : 
                        students.filter(s => String(s.room || s.roomNumber || '').trim() === currentRoomKey);
                      
                      const selectedCount = performanceSelectedStudents.length;

                      return (
                        <div className="flex-1 overflow-hidden flex flex-col bg-slate-50 relative">
                          {/* TOP BAR: Room Details Card at very top (replacing "Allocated Students" heading) */}
                          <div className="p-3 bg-emerald-50 border-b border-emerald-100 flex items-center justify-between shrink-0 shadow-xs">
                            <div className="flex items-center gap-2.5">
                              <div className="w-9 h-9 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-bold shadow-xs">
                                <School className="w-5 h-5" />
                              </div>
                              <div>
                                <span className="text-[10px] font-extrabold text-emerald-700 uppercase tracking-wider block leading-tight">{selectedHostel}</span>
                                <h3 className="text-[#1A365D] font-extrabold text-sm uppercase leading-tight">{selectedRoom} ({allocatedRoomStudents.length} Students)</h3>
                              </div>
                            </div>

                            <div className="flex items-center gap-2">
                              <button
                                onClick={() => {
                                  if (selectedCount === allocatedRoomStudents.length) {
                                    setPerformanceSelectedStudents([]);
                                  } else {
                                    setPerformanceSelectedStudents(allocatedRoomStudents.map(s => s.id));
                                  }
                                }}
                                className="text-xs font-bold text-emerald-700 hover:text-emerald-900 bg-emerald-100/80 hover:bg-emerald-200 px-3 py-1.5 rounded-xl border border-emerald-200 transition-colors"
                              >
                                {selectedCount === allocatedRoomStudents.length ? 'Deselect All' : 'Select All'}
                              </button>
                              <button
                                onClick={() => {
                                  setSelectedRoom(null);
                                  setPerformanceSelectedStudents([]);
                                }}
                                className="text-xs font-bold text-slate-600 hover:text-slate-800 bg-white hover:bg-slate-100 px-3 py-1.5 rounded-xl border border-slate-200 transition-colors shadow-2xs"
                              >
                                Change
                              </button>
                            </div>
                          </div>

                          {/* MAIN AREA: Student Checklist takes maximum height, showing all allocated students */}
                          <div className="flex-1 overflow-y-auto p-3 sm:p-4 flex flex-col gap-2.5 pb-20">
                            {allocatedRoomStudents.map(student => {
                              const isChecked = performanceSelectedStudents.includes(student.id);
                              return (
                                <button
                                  key={student.id}
                                  onClick={() => {
                                    setPerformanceSelectedStudents(prev =>
                                      prev.includes(student.id) ? prev.filter(id => id !== student.id) : [...prev, student.id]
                                    );
                                  }}
                                  className={`p-3.5 border rounded-2xl flex items-center justify-between transition-all active:scale-[0.99] text-left shadow-2xs ${
                                    isChecked
                                      ? 'bg-emerald-50/80 border-emerald-500 text-emerald-950 shadow-xs'
                                      : 'bg-white border-slate-200 text-slate-700 hover:border-slate-300 hover:bg-slate-50/50'
                                  }`}
                                >
                                  <div className="flex items-center gap-3">
                                    <div className={`w-8 h-8 rounded-xl flex items-center justify-center font-bold text-xs ${
                                      isChecked ? 'bg-emerald-600 text-white' : 'bg-slate-100 text-slate-600'
                                    }`}>
                                      {student.name.charAt(0)}
                                    </div>
                                    <div className="flex flex-col">
                                      <span className="font-extrabold text-sm leading-tight text-slate-800">{student.name}</span>
                                      <span className="text-[10px] font-bold text-slate-400 uppercase mt-0.5">
                                        Class <span className="text-emerald-700 font-extrabold">{student.class}</span>
                                      </span>
                                    </div>
                                  </div>
                                  <div className={`w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-colors ${
                                    isChecked ? 'bg-emerald-600 border-emerald-600' : 'border-slate-300 bg-white'
                                  }`}>
                                    {isChecked && <Check className="w-3.5 h-3.5 text-white stroke-[3]" />}
                                  </div>
                                </button>
                              );
                            })}

                            {allocatedRoomStudents.length === 0 && (
                              <div className="p-8 text-center bg-white rounded-2xl border border-dashed border-slate-200 text-slate-400 font-bold text-xs">
                                No students allocated to {selectedRoom}.
                              </div>
                            )}
                          </div>

                          {/* BOTTOM FIXED BAR: Compact side-by-side Tally & Fine action buttons fixed at the bottom */}
                          <div className="absolute bottom-0 left-0 right-0 p-3 bg-white/95 backdrop-blur-md border-t border-slate-200 flex items-center gap-3 z-20 shadow-lg">
                            {/* 1. TALLY BUTTON */}
                            <button
                              onClick={() => {
                                if (selectedCount === 0) {
                                  alert('Please select at least one student from the room first!');
                                  return;
                                }
                                setRoomTallyCount(1);
                                setRoomTallyReason('Room untidy');
                                setShowRoomTallyModal(true);
                              }}
                              className="flex-1 py-3 px-4 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white rounded-xl font-extrabold text-xs shadow-md hover:shadow-lg transition-all active:scale-[0.98] flex items-center justify-center gap-2 min-h-[44px]"
                            >
                              <Award className="w-4 h-4 shrink-0" />
                              <span className="truncate">1. Tally ({selectedCount})</span>
                            </button>

                            {/* 2. FINE BUTTON */}
                            <button
                              onClick={() => {
                                if (selectedCount === 0) {
                                  alert('Please select at least one student from the room first!');
                                  return;
                                }
                                setRoomFineAmount('50');
                                setRoomFineReason('Room untidy / Disturbance');
                                setShowRoomFineModal(true);
                              }}
                              className="flex-1 py-3 px-4 bg-gradient-to-r from-amber-500 to-orange-600 hover:from-amber-600 hover:to-orange-700 text-white rounded-xl font-extrabold text-xs shadow-md hover:shadow-lg transition-all active:scale-[0.98] flex items-center justify-center gap-2 min-h-[44px]"
                            >
                              <AlertTriangle className="w-4 h-4 shrink-0" />
                              <span className="truncate">2. Fine ({selectedCount})</span>
                            </button>
                          </div>
                        </div>
                      );
                    })()
                  )
                ) : !performanceSelectedClass ? (
                  <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-3">
                    {visibleClasses.map(cls => (
                      <button
                        key={cls}
                        onClick={() => setPerformanceSelectedClass(cls)}
                        className="p-4 bg-white border border-slate-200 rounded-xl shadow-sm flex items-center justify-between text-[#1A365D] hover:border-[#1A365D] transition-colors active:scale-[0.98]"
                      >
                        <span className="font-extrabold text-base uppercase">Class {cls}</span>
                        <ChevronRight className="w-5 h-5 opacity-50" />
                      </button>
                    ))}
                  </div>
                ) : (
                  <div className="flex-1 overflow-hidden flex flex-col">
                    <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-2">
                      {students.filter(s => s.class === performanceSelectedClass).map(student => (
                        <button
                          key={student.id}
                          onClick={() => {
                            setPerformanceSelectedStudents(prev => 
                              prev.includes(student.id) ? prev.filter(id => id !== student.id) : [...prev, student.id]
                            );
                          }}
                          className={`p-3 border rounded-xl flex items-center justify-between transition-colors ${
                            performanceSelectedStudents.includes(student.id)
                              ? 'bg-[#1A365D]/10 border-[#1A365D] text-[#1A365D]'
                              : 'bg-white border-slate-200 text-slate-700 hover:border-[#1A365D]/50'
                          }`}
                        >
                          <span className="font-bold text-sm">{student.name}</span>
                          <div className={`w-5 h-5 rounded-md border-2 flex items-center justify-center ${
                            performanceSelectedStudents.includes(student.id) ? 'bg-[#1A365D] border-[#1A365D]' : 'border-slate-300'
                          }`}>
                            {performanceSelectedStudents.includes(student.id) && <Check className="w-3 h-3 text-white" />}
                          </div>
                        </button>
                      ))}
                      {students.filter(s => s.class === performanceSelectedClass).length === 0 && (
                        <div className="p-6 text-center text-slate-400 font-bold text-sm">No students in this class.</div>
                      )}
                    </div>
                    <div className="p-4 bg-white border-t border-slate-200 shrink-0">
                      <button
                        onClick={() => setShowPerformanceSubmitModal(true)}
                        disabled={performanceSelectedStudents.length === 0}
                        className="w-full py-3.5 px-4 bg-[#1A365D] hover:bg-[#2A4365] disabled:bg-slate-300 text-white rounded-xl font-extrabold text-sm shadow-md transition-colors active:scale-[0.98]"
                      >
                        Submit Selected ({performanceSelectedStudents.length})
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        ) : (
          /* SPREADSHEET VIEW (RBAC Protected) */
          <div className="flex-1 flex flex-col overflow-hidden bg-[#F8F9FA]">
            {/* Top Sub-tabs & Settings Bar (Admin Only) */}
            {isAdminAuthenticated && (
              <div className="flex items-center gap-1 p-2 bg-white border-b border-slate-200 shrink-0">
                <button
                  onClick={() => setAdminSubTab('sheet')}
                  className={`flex-1 py-2 rounded-xl text-xs font-extrabold uppercase tracking-wider transition-all ${adminSubTab === 'sheet' ? 'bg-[#1A365D] text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100'}`}
                >Score Sheet</button>
                {isSuperAdmin && (
                  <button
                    onClick={() => { setAdminSubTab('users'); fetchUsers(); }}
                    className={`flex-1 py-2 rounded-xl text-xs font-extrabold uppercase tracking-wider transition-all ${adminSubTab === 'users' ? 'bg-[#1A365D] text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100'}`}
                  >Users</button>
                )}
                <button
                  onClick={() => setShowAdminSettingsModal(true)}
                  className="p-2 rounded-xl hover:bg-slate-100 text-slate-400 hover:text-slate-600 transition-all"
                  title="Admin Settings"
                ><Settings className="w-4 h-4" /></button>
                <button
                  onClick={handleLogout}
                  className="p-2 rounded-xl hover:bg-red-50 text-slate-400 hover:text-red-500 transition-all"
                  title="Logout"
                ><Lock className="w-4 h-4" /></button>
              </div>
            )}

              {isSuperAdmin && adminSubTab === 'users' ? (
                /* ─── USERS MANAGEMENT UI ─── */
                <div className="flex-1 flex flex-col overflow-hidden">
                  <div className="p-4 flex items-center justify-between bg-white border-b border-slate-100 shrink-0">
                    <div>
                      <h2 className="font-extrabold text-[#1A365D] text-sm uppercase tracking-wider">User Management</h2>
                      <p className="text-xs text-slate-400 mt-0.5">{usersList.length} users registered</p>
                    </div>
                    <button
                      onClick={() => { setEditingUser(null); setUserForm({ username: '', password: '', role: 'user', permissions: [] }); setUserFormError(''); setShowUserModal(true); }}
                      className="flex items-center gap-1.5 px-3 py-2 bg-[#1A365D] text-white text-xs font-bold rounded-xl hover:bg-[#2A4365] transition-colors shadow-sm"
                    >
                      <PlusCircle className="w-3.5 h-3.5" /> Add User
                    </button>
                  </div>
                  <div className="flex-1 overflow-y-auto p-3 flex flex-col gap-2">
                    {usersLoading ? (
                      <div className="flex items-center justify-center py-12"><div className="w-8 h-8 border-4 border-[#1A365D] border-t-transparent rounded-full animate-spin" /></div>
                    ) : usersList.length === 0 ? (
                      <div className="text-center py-12 text-slate-400 text-sm font-bold">No users found.</div>
                    ) : usersList.map(u => (
                      <div key={u.id} className="bg-white border border-slate-200 rounded-xl p-3.5 flex items-start gap-3 shadow-sm">
                        <div className="w-9 h-9 rounded-xl bg-[#1A365D]/10 flex items-center justify-center text-[#1A365D] shrink-0 font-black text-sm">
                          {u.username[0].toUpperCase()}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2">
                            <span className="font-extrabold text-slate-800 text-sm">{u.username}</span>
                            <span className={`text-[10px] font-extrabold px-2 py-0.5 rounded-full uppercase ${
                              u.role === 'super_admin' ? 'bg-amber-100 text-amber-800 border border-amber-300' :
                              u.role === 'admin' ? 'bg-purple-100 text-purple-700' : 'bg-slate-100 text-slate-600'
                            }`}>
                              {u.role === 'super_admin' ? 'Super Admin' : u.role}
                            </span>
                          </div>
                          <div className="flex flex-wrap gap-1 mt-1.5">
                            {u.role === 'super_admin' ? (
                              <span className="text-[10px] bg-amber-50 text-amber-700 border border-amber-200 rounded-full px-2 py-0.5 font-bold">Super Admin (User Control + All Access)</span>
                            ) : u.role === 'admin' ? (
                              <span className="text-[10px] bg-purple-50 text-purple-600 border border-purple-200 rounded-full px-2 py-0.5 font-bold">Full Features Access</span>
                            ) : (Array.isArray(u.permissions) ? u.permissions : []).map(p => (
                              <span key={p} className="text-[10px] bg-blue-50 text-blue-600 border border-blue-200 rounded-full px-2 py-0.5 font-semibold">{p}</span>
                            ))}
                            {(u.role !== 'admin' && u.role !== 'super_admin' && (!u.permissions || u.permissions.length === 0)) && (
                              <span className="text-[10px] text-slate-400 italic">No permissions</span>
                            )}
                          </div>
                        </div>
                        <div className="flex gap-1 shrink-0">
                          <button
                            onClick={() => { setEditingUser(u); setUserForm({ username: u.username, password: '', role: u.role, permissions: Array.isArray(u.permissions) ? u.permissions : [] }); setUserFormError(''); setShowUserModal(true); }}
                            className="p-1.5 rounded-lg hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors"
                          ><Pencil className="w-3.5 h-3.5" /></button>
                          {u.username !== 'admin' && (
                            <button
                              onClick={() => handleDeleteUser(u.id)}
                              className="p-1.5 rounded-lg hover:bg-red-50 text-slate-400 hover:text-red-500 transition-colors"
                            ><Trash2 className="w-3.5 h-3.5" /></button>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ) : (
              /* ─── SCORE SHEET ─── */
              <div className="flex-1 flex flex-col overflow-hidden p-4 gap-3.5">
              <div className="flex justify-between items-center pb-2 border-b border-slate-200">
                <div className="flex items-center gap-2 text-slate-700">
                  <Table className="text-[#1A365D] w-4.5 h-4.5" />
                  <span className="text-xs font-extrabold uppercase tracking-wider text-[#1A365D]">Score Sheet</span>
                </div>

              {/* Class Dropdown */}
              <select
                value={adminClass}
                onChange={e => setAdminClass(e.target.value)}
                className="bg-white border border-slate-200 text-[#1A365D] text-xs font-extrabold uppercase py-1.5 px-3 rounded-lg focus:outline-none focus:border-[#1A365D] shadow-sm cursor-pointer"
              >
                {visibleClasses.map(cls => (
                  <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>
                ))}
              </select>
            </div>

            {/* Score calculations rule banner */}

            <div className="text-[11px] text-slate-600 bg-white border border-slate-200 p-2.5 px-3.5 rounded-xl shadow-xs flex items-center justify-between">
              <span>Rule: <strong className="text-amber-700">1 Star = +2 Marks</strong> | <strong className="text-sky-700">1 Tally = -1 Mark</strong></span>
              <span className="text-[#1A365D] font-extrabold uppercase text-[9px] bg-[#1A365D]/10 border border-[#1A365D]/20 px-2 py-0.5 rounded-full">
                Active
              </span>
            </div>

            {/* Action Buttons (RBAC Filtered) */}
            <div className={`grid ${isAdminAuthenticated ? 'grid-cols-4' : 'grid-cols-1'} gap-2`}>
              <button
                onClick={() => {
                  setDownloadSelectedClasses(CLASSES);
                  setDateModalNextAction('ADMIN_REPORT');
                  setShowIRDateModal(true);
                }}
                className="flex items-center justify-center gap-1.5 py-2.5 px-1 rounded-xl font-extrabold text-[11px] bg-emerald-700 hover:bg-emerald-800 text-white shadow-xs active:scale-[0.98] transition-all"
              >
                <Download className="w-3.5 h-3.5 shrink-0" />
                Report
              </button>
              
              {isAdminAuthenticated && (
                <>
                  <button
                    onClick={() => {
                      setAddStudentMethod('single');
                      setShowAddModal(true);
                    }}
                    className="flex items-center justify-center gap-1.5 py-2.5 px-1 rounded-xl font-extrabold text-[11px] bg-[#1A365D] hover:bg-[#2A4365] text-white shadow-xs active:scale-[0.98] transition-all"
                  >
                    <PlusCircle className="w-3.5 h-3.5 shrink-0" />
                    Add Student
                  </button>

                  <button
                    onClick={handleAdminDeleteSelectedStudents}
                    className="flex items-center justify-center gap-1.5 py-2.5 px-1 rounded-xl font-extrabold text-[11px] bg-amber-500 hover:bg-amber-600 text-white shadow-xs active:scale-[0.98] transition-all"
                  >
                    <Trash2 className="w-3.5 h-3.5 shrink-0" />
                    Delete
                  </button>

                  <button
                    onClick={() => {
                      setClearSelectedClasses([]);
                      setShowClearModal(true);
                    }}
                    className="flex items-center justify-center gap-1.5 py-2.5 px-1 rounded-xl font-extrabold text-[11px] bg-rose-600 hover:bg-rose-700 text-white shadow-xs active:scale-[0.98] transition-all"
                  >
                    <Trash2 className="w-3.5 h-3.5 shrink-0" />
                    Clear Data
                  </button>
                </>
              )}
            </div>

            {/* Spreadsheet Table */}
            <div className="flex-1 overflow-auto border border-slate-200 rounded-xl bg-white shadow-sm">
              <table className="w-full border-collapse text-left text-xs font-sans">
                <thead>
                  {/* Outer Class Header Row */}
                  <tr className="border-b border-slate-200 bg-slate-100 text-center font-bold">
                    <th className="border-r border-slate-200 p-2 text-slate-400 font-mono text-[10px] w-8">
                      <input 
                        type="checkbox" 
                        className="cursor-pointer"
                        checked={adminClassStudents.length > 0 && adminSelectedStudents.length === adminClassStudents.length}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setAdminSelectedStudents(adminClassStudents.map(s => s.id));
                          } else {
                            setAdminSelectedStudents([]);
                          }
                        }}
                      />
                    </th>
                    <th className="border-r border-slate-200 p-2 text-slate-400 font-mono text-[10px] w-8">#</th>
                    <th colSpan="12" className="p-2 text-[#1A365D] font-extrabold uppercase tracking-widest text-xs bg-[#1A365D]/5">
                      Class {adminClass.toUpperCase()}
                    </th>
                  </tr>
                  
                  {/* Table headers */}
                  <tr className="border-b border-slate-200 bg-slate-50 text-slate-500 uppercase text-[10px]">
                    <th className="border-r border-slate-200 p-2 text-center"></th>
                    <th className="border-r border-slate-200 p-2 text-center font-mono text-[10px]">ROW</th>
                    <th className="border-r border-slate-200 p-2.5 font-bold">STUDENT NAME</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-amber-700">STARS</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-sky-700">TALLIES</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-[#1A365D]">TOTAL</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-purple-700">GRADE</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-orange-600">N&O TALLY</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-[#1A365D]">TOTAL</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-orange-700">N&O GRADE</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-sky-600">DIARY TALLIES</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-cyan-600">SHEETS</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-rose-600">FINE</th>
                    <th className="border-r border-slate-200 p-2.5 text-center font-bold text-[#1A365D]">TOTAL</th>
                    <th className="p-2.5 text-center font-bold text-indigo-700">ATTITUDE GRADE</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {adminClassStudents.length > 0 ? (
                    adminClassStudents.map((student, idx) => {
                      const computedTotal = calculateTotalScore(student.star, student.tally);
                      const finalTotal = (student.customTotal !== undefined && student.customTotal !== null && student.customTotal !== '')
                        ? Number(student.customTotal)
                        : computedTotal;
                      const computedGrade = calculateGrade(finalTotal);
                      const finalGrade = (student.customGrade && student.customGrade.trim())
                        ? student.customGrade.trim().toUpperCase()
                        : computedGrade;
                      const attitudeTotal = ((Number(student.diaryTally) || 0) * -0.5) + (getFineCount(student) * -1.5) + (Number(student.sheetTally) || 0);
                      return (
                        <tr key={student.id} className="hover:bg-slate-50 transition-colors">
                          <td className="border-r border-slate-200 p-2 text-center text-slate-400">
                            <input 
                              type="checkbox" 
                              className="cursor-pointer"
                              checked={adminSelectedStudents.includes(student.id)}
                              onChange={(e) => {
                                if (e.target.checked) {
                                  setAdminSelectedStudents(prev => [...prev, student.id]);
                                } else {
                                  setAdminSelectedStudents(prev => prev.filter(id => id !== student.id));
                                }
                              }}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-2 text-center text-slate-400 font-mono select-none text-[10px]">
                            {idx + 1}
                          </td>
                          <td className="border-r border-slate-200 p-1 font-semibold text-slate-800">
                            <input
                              type="text"
                              value={student.name}
                              readOnly={!isAdminAuthenticated}
                              onChange={(e) => updateStudentField(student.id, 'name', e.target.value)}
                              className={`bg-transparent text-left w-full focus:outline-none py-1 px-2 rounded font-semibold text-xs text-slate-800 ${isAdminAuthenticated ? 'focus:bg-slate-100 cursor-pointer' : 'cursor-default'}`}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-1 text-center text-amber-700 font-bold">
                            <input
                              type="number"
                              min="0"
                              readOnly={!isAdminAuthenticated}
                              value={(student.star || 0) + (student.diaryStar || 0)}
                              onChange={(e) => {
                                if (!isAdminAuthenticated) return;
                                const val = Math.max(0, parseInt(e.target.value) || 0);
                                setStudents(prev => prev.map(s => s.id === student.id ? { ...s, star: val - (s.diaryStar || 0), customTotal: null, customGrade: '' } : s));
                                debounceUpdateStudent(student.id);
                              }}
                              className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-bold text-amber-700 ${isAdminAuthenticated ? 'focus:bg-amber-50 cursor-pointer' : 'cursor-default'}`}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-1 text-center text-sky-700 font-bold">
                            <input
                              type="number"
                              min="0"
                              readOnly={!isAdminAuthenticated}
                              value={student.tally}
                              onChange={(e) => {
                                if (!isAdminAuthenticated) return;
                                const val = Math.max(0, parseInt(e.target.value) || 0);
                                setStudents(prev => prev.map(s => s.id === student.id ? { ...s, tally: val, customTotal: null, customGrade: '' } : s));
                                debounceUpdateStudent(student.id);
                              }}
                              className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-bold text-sky-700 ${isAdminAuthenticated ? 'focus:bg-sky-50 cursor-pointer' : 'cursor-default'}`}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-1 text-center font-extrabold">
                            <input
                              type="number"
                              readOnly={!isAdminAuthenticated}
                              value={finalTotal}
                              onChange={(e) => {
                                if (!isAdminAuthenticated) return;
                                const raw = e.target.value;
                                if (raw === '') {
                                  setStudents(prev => prev.map(s => s.id === student.id ? { ...s, customTotal: null, customGrade: '' } : s));
                                  debounceUpdateStudent(student.id);
                                } else {
                                  const val = parseInt(raw, 10);
                                  if (!isNaN(val)) {
                                    setStudents(prev => prev.map(s => s.id === student.id ? { ...s, customTotal: val, customGrade: '' } : s));
                                    debounceUpdateStudent(student.id);
                                  }
                                }
                              }}
                              className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-extrabold ${
                                finalTotal > 0 ? 'text-emerald-700' : finalTotal < 0 ? 'text-rose-600' : 'text-slate-500'
                              } ${isAdminAuthenticated ? 'focus:bg-slate-100 cursor-pointer' : 'cursor-default'}`}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-1 text-center font-extrabold">
                            <input
                              type="text"
                              readOnly={!isAdminAuthenticated}
                              value={finalGrade}
                              onChange={(e) => {
                                if (!isAdminAuthenticated) return;
                                updateStudentField(student.id, 'customGrade', e.target.value.toUpperCase());
                              }}
                              className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-extrabold uppercase ${
                                finalTotal >= 20 ? 'text-emerald-700' : finalTotal >= 7 ? 'text-emerald-500' : finalTotal >= 0 ? 'text-amber-500' : finalTotal >= -6 ? 'text-orange-500' : finalTotal >= -20 ? 'text-rose-500' : 'text-rose-700'
                              } ${isAdminAuthenticated ? 'focus:bg-purple-50 cursor-pointer' : 'cursor-default'}`}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-1 text-center font-bold text-orange-600">
                            <input
                              type="number"
                              min="0"
                              readOnly={!isAdminAuthenticated}
                              value={student.neatAndOrderTally || 0}
                              onChange={(e) => {
                                if (!isAdminAuthenticated) return;
                                const val = Math.max(0, parseInt(e.target.value) || 0);
                                updateStudentField(student.id, 'neatAndOrderTally', val);
                              }}
                              className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-bold text-orange-600 ${isAdminAuthenticated ? 'focus:bg-orange-50 cursor-pointer' : 'cursor-default'}`}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-2.5 text-center font-extrabold text-[#1A365D]">
                            {-getNOIncidents(student)}
                          </td>
                          <td className="border-r border-slate-200 p-2.5 text-center font-semibold text-orange-700 uppercase">
                            {calculateNOGrade(-getNOIncidents(student))}
                          </td>
                          <td className="border-r border-slate-200 p-1 text-center text-sky-500 font-bold">
                            <input
                              type="number"
                              min="0"
                              readOnly={!isAdminAuthenticated}
                              value={student.diaryTally || 0}
                              onChange={(e) => {
                                if (!isAdminAuthenticated) return;
                                const val = Math.max(0, parseInt(e.target.value) || 0);
                                updateStudentField(student.id, 'diaryTally', val);
                              }}
                              className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-bold text-sky-600 ${isAdminAuthenticated ? 'focus:bg-sky-50 cursor-pointer' : 'cursor-default'}`}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-1 text-center text-cyan-500 font-bold">
                            <input
                              type="number"
                              readOnly={!isAdminAuthenticated}
                              value={student.sheetTally || 0}
                              onChange={(e) => {
                                if (!isAdminAuthenticated) return;
                                updateStudentField(student.id, 'sheetTally', parseInt(e.target.value) || 0);
                              }}
                              className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-bold text-cyan-600 ${isAdminAuthenticated ? 'focus:bg-cyan-50 cursor-pointer' : 'cursor-default'}`}
                            />
                          </td>
                          <td className="border-r border-slate-200 p-1 text-center font-bold text-rose-600">
                            <input
                              type="number"
                              min="0"
                              value={student.fine || 0}
                              onChange={(e) => {
                                const val = Math.max(0, parseInt(e.target.value) || 0);
                                updateStudentField(student.id, 'fine', val);
                              }}
                              className="bg-transparent text-center w-full focus:outline-none focus:bg-rose-50 py-1 px-1 rounded font-bold text-rose-600"
                            />
                          </td>
                          <td className={`border-r border-slate-200 p-2.5 text-center font-extrabold text-[#1A365D]`}>
                            {attitudeTotal}
                          </td>
                          <td className="p-1 font-semibold text-indigo-700">
                            <input
                              type="text"
                              value={calculateAttitudeGrade(attitudeTotal)}
                              readOnly={true}
                              className="bg-transparent text-center w-full focus:outline-none focus:bg-indigo-50 py-1 px-1 rounded font-semibold text-xs text-indigo-700 uppercase"
                            />
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan="14" className="p-8 text-center text-slate-400 italic">
                        No students in this class.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
          )}
        </div>
      )}
    </div>







      {/* User Add/Edit Modal */}
      {showUserModal && (
        <div className="fixed inset-0 bg-slate-900/70 backdrop-blur-sm z-50 flex items-end justify-center">
          <div className="w-full max-w-md bg-white rounded-t-3xl shadow-2xl p-6 flex flex-col gap-4 animate-[modalSlideUp_0.25s_ease-out] max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between">
              <h3 className="text-base font-extrabold text-[#1A365D]">{editingUser ? 'Edit User' : 'Add New User'}</h3>
              <button onClick={() => { setShowUserModal(false); setUserFormError(''); }} className="p-2 rounded-xl hover:bg-slate-100 text-slate-400 hover:text-slate-600 transition-colors"><X className="w-4 h-4" /></button>
            </div>
            <form onSubmit={handleSaveUser} className="flex flex-col gap-4">
              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">Username</label>
                <input type="text" value={userForm.username} onChange={e => setUserForm(f => ({ ...f, username: e.target.value }))}
                  className="mt-1.5 w-full px-3 py-2.5 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50 transition-colors"
                  placeholder="e.g. teacher1" required />
              </div>
              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">Password {editingUser && <span className="font-normal text-slate-400">(leave blank to keep current)</span>}</label>
                <input type="password" value={userForm.password} onChange={e => setUserForm(f => ({ ...f, password: e.target.value }))}
                  className="mt-1.5 w-full px-3 py-2.5 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50 transition-colors"
                  placeholder={editingUser ? "Enter new password..." : "Enter password..."} />
              </div>
              <div>
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">Role</label>
                <select value={userForm.role} onChange={e => setUserForm(f => ({ ...f, role: e.target.value }))}
                  className="mt-1.5 w-full px-3 py-2.5 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50 transition-colors cursor-pointer">
                  <option value="user">User</option>
                  <option value="admin">Admin</option>
                  <option value="super_admin">Super Admin</option>
                </select>
              </div>
              {userForm.role !== 'admin' && userForm.role !== 'super_admin' && (
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 block">Permissions</label>
                  <div className="grid grid-cols-2 gap-2">
                    {ALL_PERMISSIONS.map(perm => (
                      <label key={perm.key} className={`flex items-center gap-2 p-2.5 rounded-xl border cursor-pointer transition-all ${userForm.permissions.includes(perm.key) ? 'border-[#1A365D] bg-[#1A365D]/5' : 'border-slate-200 hover:border-slate-300 bg-white'}`}>
                        <input type="checkbox" checked={userForm.permissions.includes(perm.key)}
                          onChange={e => setUserForm(f => ({ ...f, permissions: e.target.checked ? [...f.permissions, perm.key] : f.permissions.filter(p => p !== perm.key) }))}
                          className="w-4 h-4 accent-[#1A365D]" />
                        <span className="text-xs font-semibold text-slate-700">{perm.label}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}
              {userFormError && <p className="text-xs font-bold text-red-600 bg-red-50 border border-red-200 rounded-xl px-3 py-2">{userFormError}</p>}
              <div className="flex gap-2 mt-1">
                <button type="button" onClick={() => { setShowUserModal(false); setUserFormError(''); }}
                  className="flex-1 py-3 rounded-xl font-bold text-sm border border-slate-200 hover:bg-slate-50 text-slate-600 transition-colors">Cancel</button>
                <button type="submit"
                  className="flex-1 py-3 rounded-xl font-extrabold text-sm bg-[#1A365D] hover:bg-[#2A4365] text-white shadow-md transition-colors active:scale-95">{editingUser ? 'Save Changes' : 'Create User'}</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Persistent Bottom Tab Navigation Bar (4 Tabs: Scoring, Mentor, Performance, Admin Sheet) */}
      <footer className="h-16 border-t border-slate-200 bg-white flex items-center justify-around px-2 shadow-[0_-4px_16px_rgba(0,0,0,0.03)] shrink-0 z-10 font-sans">
        <button
          onClick={() => {
            setActiveTab('scoring');
            setSelectedClass(null);
            setSelectedStudentIds([]);
            setIsScoring(false);
            setSearchQuery('');
          }}
          className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 ${
            activeTab === 'scoring' 
              ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]' 
              : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
          }`}
        >
          <Sparkles className={`w-4 h-4 ${activeTab === 'scoring' ? 'text-amber-300' : ''}`} />
          <span className="text-[10px] tracking-wide">Scoring</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('mentor');
            setSaveStatus('');
          }}
          className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 ${
            activeTab === 'mentor' 
              ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]' 
              : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
          }`}
        >
          <UserCheck className={`w-4 h-4 ${activeTab === 'mentor' ? 'text-sky-300' : ''}`} />
          <span className="text-[10px] tracking-wide">Mentor</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('performance');
            setSaveStatus('');
          }}
          className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 ${
            activeTab === 'performance' 
              ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]' 
              : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
          }`}
        >
          <TrendingUp className={`w-4 h-4 ${activeTab === 'performance' ? 'text-emerald-300' : ''}`} />
          <span className="text-[10px] tracking-wide">Performance</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('admin');
            setSaveStatus('');
          }}
          className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 ${
            activeTab === 'admin' 
              ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]' 
              : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
          }`}
        >
          <Table className={`w-4 h-4 ${activeTab === 'admin' ? 'text-purple-300' : ''}`} />
          <span className="text-[10px] tracking-wide">
            {isAdminAuthenticated ? 'Admin Sheet' : 'Score Sheet'}
          </span>
        </button>
      </footer>

      {/* Summary Share & Return Modal */}
      {showSummaryModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-[360px] bg-white border border-slate-200 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 animate-[modalSlideUp_0.2s_ease-out]">
            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
              <h3 className="text-base font-extrabold text-[#333333] flex items-center gap-2">
                <Check className="w-5 h-5 text-emerald-600" />
                Evaluation Summary
              </h3>
            </div>

            <div className="flex flex-col gap-2">
              <p className="text-[10px] text-slate-500 font-bold uppercase tracking-wider">
                Generated Report Text
              </p>
              
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-xs text-slate-800 font-mono whitespace-pre-wrap select-all max-h-48 overflow-y-auto">
                {summaryText}
              </div>
            </div>

            {/* Actions for Sharing */}
            <div className="flex flex-col gap-2 pt-2 border-t border-slate-100">
              <div className="grid grid-cols-2 gap-2">
                <button
                  onClick={handleCopyText}
                  className="flex items-center justify-center gap-1.5 py-2.5 px-3 rounded-xl bg-slate-100 border border-slate-200 text-xs font-bold text-slate-700 hover:bg-slate-200 active:scale-[0.98] transition-all"
                >
                  <Copy className="w-3.5 h-3.5 text-slate-600" />
                  {copyFeedback ? copyFeedback : 'Copy Text'}
                </button>
                <a
                  href={`https://api.whatsapp.com/send?text=${encodeURIComponent(summaryText)}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="flex items-center justify-center gap-1.5 py-2.5 px-3 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-xs font-bold text-white text-center active:scale-[0.98] transition-all shadow-xs"
                >
                  <Share2 className="w-3.5 h-3.5" />
                  WhatsApp
                </a>
              </div>

              <button
                onClick={handleCloseSummaryAndGoHome}
                className="w-full mt-1 py-3 rounded-xl bg-[#1A365D] text-white text-xs font-extrabold shadow-md hover:bg-[#2A4365] active:scale-[0.98] transition-all"
              >
                Done & Return to Overview
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Add Student Modal */}
      {showAddModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-[360px] bg-white border border-slate-200 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 animate-[modalSlideUp_0.2s_ease-out]">
            <div className="flex justify-between items-center pb-2.5 border-b border-slate-100">
              <h3 className="text-base font-extrabold text-[#333333] flex items-center gap-2">
                <PlusCircle className="w-5 h-5 text-[#1A365D]" />
                Add Student
              </h3>
              <button 
                type="button"
                onClick={() => {
                  setShowAddModal(false);
                  setNewStudentName('');
                  setAddStudentMethod('single');
                }}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Tab selector */}
            <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200">
              <button
                type="button"
                onClick={() => setAddStudentMethod('single')}
                className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${
                  addStudentMethod === 'single'
                    ? 'bg-[#1A365D] text-white shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Add Single Student
              </button>
              <button
                type="button"
                onClick={() => setAddStudentMethod('excel')}
                className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${
                  addStudentMethod === 'excel'
                    ? 'bg-[#1A365D] text-white shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Excel Roster Import
              </button>
            </div>

            {addStudentMethod === 'single' ? (
              <form onSubmit={handleAddStudent} className="flex flex-col gap-4">
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                    Student Full Name
                  </label>
                  <input
                    type="text"
                    autoFocus
                    required
                    value={newStudentName}
                    onChange={e => setNewStudentName(e.target.value)}
                    placeholder="Enter student full name..."
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all font-semibold"
                  />
                  <span className="text-[11px] text-slate-500 font-medium">
                    Enrolling into <strong className="uppercase text-[#1A365D]">Class {activeTab === 'admin' ? adminClass : selectedClass}</strong>.
                  </span>
                </div>

                <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => {
                      setShowAddModal(false);
                      setNewStudentName('');
                    }}
                    className="px-4 py-2.5 rounded-xl font-bold text-xs border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 transition-all"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-4 py-2.5 rounded-xl font-extrabold text-xs bg-[#1A365D] text-white hover:bg-[#2A4365] active:scale-[0.98] transition-all shadow-md"
                  >
                    Save Student
                  </button>
                </div>
              </form>
            ) : (
              <div className="flex flex-col gap-4">
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                    Roster Excel File (.xlsx, .xls, .csv)
                  </label>
                  
                  <div className="text-[11px] text-slate-600 leading-relaxed bg-slate-50 p-3 rounded-xl border border-slate-200 flex flex-col gap-1">
                    <span className="font-extrabold text-[#1A365D]">Excel File Guidelines:</span>
                    <span>1. Sheet names should match class codes (e.g. <code>S2B</code>).</span>
                    <span>2. Columns must have headers <code>rol</code> and <code>name</code>.</span>
                  </div>

                  <div className="bg-amber-50 border border-amber-200 p-3 rounded-xl text-[11px] text-amber-800 font-medium leading-normal flex items-start gap-2 mt-1">
                    <span className="font-extrabold shrink-0">⚠️ Note:</span>
                    <span>Importing will update the roster for the matching classroom sheets.</span>
                  </div>
                </div>

                <div className="flex flex-col gap-2 pt-2 border-t border-slate-100">
                  <label
                    className="w-full flex items-center justify-center gap-2 py-3 px-4 rounded-xl font-extrabold text-xs bg-emerald-700 hover:bg-emerald-800 text-white shadow-md active:scale-[0.98] transition-all cursor-pointer text-center"
                  >
                    <FileText className="w-4 h-4 text-white" />
                    Browse & Import Excel Roster
                    <input 
                      type="file" 
                      accept=".xlsx,.xls,.csv" 
                      onChange={handleImportExcel} 
                      className="hidden" 
                    />
                  </label>
                  
                  <button
                    type="button"
                    onClick={() => setShowAddModal(false)}
                    className="w-full py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-600 hover:text-slate-800 text-xs font-bold transition-all"
                  >
                    Cancel
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Bulk Text Entry Modal */}
      {showBulkModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <form 
            onSubmit={handleBulkSubmit}
            className="w-full max-w-[380px] bg-white border border-slate-200 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 animate-[modalSlideUp_0.2s_ease-out]"
          >
            <div className="flex justify-between items-center pb-2 border-b border-slate-100">
              <h3 className="text-base font-extrabold text-[#333333] flex items-center gap-2">
                <FileText className="w-5 h-5 text-[#1A365D]" />
                Bulk Score Text Entry
              </h3>
              <button 
                type="button"
                onClick={() => {
                  setShowBulkModal(false);
                  setBulkInputText('');
                }}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="flex flex-col gap-1.5">
              <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                Paste or type student score instructions
              </label>
              <textarea
                autoFocus
                value={bulkInputText}
                onChange={e => setBulkInputText(e.target.value)}
                placeholder="Example: Good Participation Aamir 2 Star Homework Jibin 3 Tally"
                rows={4}
                className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all font-sans resize-none"
              />
              <span className="text-[11px] text-slate-500 font-medium">
                Format: <code>[Reason] [Name] [Count] [Star/Tally]</code>
              </span>
            </div>

            {/* Dynamic Parser Preview */}
            {bulkInputText.trim() && (
              <div className="flex flex-col gap-2 border-t border-slate-100 pt-3 max-h-[150px] overflow-y-auto">
                <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                  Parsed Output Preview:
                </span>
                <div className="flex flex-col gap-1.5">
                  {parseBulkText(bulkInputText, students, selectedClass).map((entry, idx) => (
                    <div 
                      key={idx} 
                      className="flex justify-between items-center bg-slate-50 border border-slate-200 p-2.5 rounded-xl text-xs"
                    >
                      <div className="flex flex-col">
                        <span className="font-bold text-slate-800">
                          {entry.student.name}
                          {entry.student.id ? (
                            <span className="text-[10px] font-semibold text-slate-500 ml-1.5 uppercase">
                              ({entry.student.class})
                            </span>
                          ) : (
                            <span className="text-[10px] font-bold text-[#1A365D] ml-1.5 uppercase">
                              (New in {selectedClass || 's2b'})
                            </span>
                          )}
                        </span>
                        <span className="text-[10px] text-slate-500 italic mt-0.5">
                          Reason: {entry.reason}
                        </span>
                      </div>
                      <span className={`font-mono font-extrabold px-2 py-0.5 rounded-full text-xs ${
                        entry.type === 'star' 
                          ? 'text-amber-800 bg-amber-100 border border-amber-200' 
                          : 'text-sky-800 bg-sky-100 border border-sky-200'
                      }`}>
                        +{entry.amount} {entry.type === 'star' ? 'Star' : 'Tally'}
                      </span>
                    </div>
                  ))}
                  {parseBulkText(bulkInputText, students, selectedClass).length === 0 && (
                    <span className="text-xs text-amber-700 italic font-medium">
                      No valid score instructions found. Include amount and type (e.g. 2 Star, 3 Tally).
                    </span>
                  )}
                </div>
              </div>
            )}

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => {
                  setShowBulkModal(false);
                  setBulkInputText('');
                }}
                className="px-4 py-2.5 rounded-xl font-bold text-xs border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 transition-all"
              >
                Cancel
              </button>
              <button
                type="submit"
                disabled={!bulkInputText.trim() || parseBulkText(bulkInputText, students, selectedClass).length === 0}
                className="px-4 py-2.5 rounded-xl font-extrabold text-xs bg-[#1A365D] text-white shadow-md disabled:opacity-50 disabled:pointer-events-none hover:bg-[#2A4365] active:scale-[0.98] transition-all"
              >
                Apply Batch Scores
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Export Excel Modal */}
      {showDownloadModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div 
            className="w-full max-w-[360px] bg-white border border-slate-200 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 animate-[modalSlideUp_0.2s_ease-out]"
          >
            <div className="flex justify-between items-center pb-2 border-b border-slate-100">
              <h3 className="text-base font-extrabold text-[#333333] flex items-center gap-2">
                <Download className="w-5 h-5 text-emerald-600" />
                Export Excel Report
              </h3>
              <button 
                type="button"
                onClick={() => setShowDownloadModal(false)}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="flex flex-col gap-3">
              <div className="flex justify-between items-center text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                <span>Select Classes to Include:</span>
                <button
                  type="button"
                  onClick={handleToggleAllDownloadClasses}
                  className="text-[#1A365D] hover:underline transition-colors normal-case text-xs font-bold"
                >
                  {downloadSelectedClasses.length === CLASSES.length ? 'Deselect All' : 'Select All'}
                </button>
              </div>

              {/* Class Checkbox Grid */}
              <div className="grid grid-cols-2 gap-2 bg-slate-50 p-3 rounded-xl border border-slate-200 max-h-[170px] overflow-y-auto">
                {visibleClasses.map(clsName => {
                  const isChecked = downloadSelectedClasses.includes(clsName);
                  const count = students.filter(s => s.class === clsName).length;
                  return (
                    <label 
                      key={clsName}
                      className={`flex items-center gap-2 p-2.5 rounded-lg border cursor-pointer transition-all ${
                        isChecked 
                          ? 'border-emerald-600 bg-emerald-50 text-emerald-900 font-extrabold' 
                          : 'border-slate-200 bg-white hover:bg-slate-100 text-slate-600'
                      }`}
                    >
                      <input 
                        type="checkbox"
                        checked={isChecked}
                        onChange={() => handleToggleDownloadClass(clsName)}
                        className="sr-only"
                      />
                      <div 
                        className={`w-4 h-4 rounded border flex items-center justify-center transition-all shrink-0 ${
                          isChecked 
                            ? 'bg-emerald-600 border-emerald-600 text-white' 
                            : 'border-slate-300 bg-white'
                        }`}
                      >
                        {isChecked && <Check className="w-3 h-3 stroke-[3]" />}
                      </div>
                      <span className="text-xs uppercase tracking-wider">Class {clsName}</span>
                      <span className="text-[10px] font-bold opacity-70 ml-auto">({count})</span>
                    </label>
                  );
                })}
              </div>
            </div>
            
            {/* Action Buttons */}
            <div className="flex flex-col gap-2 pt-2 border-t border-slate-100">
              <button
                onClick={handleDownloadExcel}
                disabled={downloadSelectedClasses.length === 0}
                className="w-full flex items-center justify-center gap-2 py-3 px-4 rounded-xl font-extrabold text-xs bg-emerald-700 text-white hover:bg-emerald-800 disabled:opacity-50 disabled:cursor-not-allowed active:scale-[0.98] transition-all shadow-md"
              >
                <Download className="w-4 h-4" />
                Download Report ({downloadSelectedClasses.length})
              </button>
              
              <button
                onClick={() => setShowDownloadModal(false)}
                className="w-full py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-600 hover:text-slate-800 text-xs font-bold transition-all"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Clear Data Modal */}
      {showClearModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div 
            className="w-full max-w-[360px] bg-white border border-slate-200 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 animate-[modalSlideUp_0.2s_ease-out]"
          >
            <div className="flex justify-between items-center pb-2 border-b border-slate-100">
              <h3 className="text-base font-extrabold text-[#333333] flex items-center gap-2">
                <Trash2 className="w-5 h-5 text-rose-600" />
                Clear Classroom Data
              </h3>
              <button 
                type="button"
                onClick={() => setShowClearModal(false)}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="flex flex-col gap-3">
              <div className="flex justify-between items-center text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                <span>Select Classes to Clear:</span>
                <button
                  type="button"
                  onClick={handleToggleAllClearClasses}
                  className="text-rose-600 hover:underline transition-colors normal-case text-xs font-bold"
                >
                  {clearSelectedClasses.length === CLASSES.length ? 'Deselect All' : 'Select All'}
                </button>
              </div>

              {/* Class Checkbox Grid */}
              <div className="grid grid-cols-2 gap-2 bg-slate-50 p-3 rounded-xl border border-slate-200 max-h-[170px] overflow-y-auto">
                {visibleClasses.map(clsName => {
                  const isChecked = clearSelectedClasses.includes(clsName);
                  const count = students.filter(s => s.class === clsName).length;
                  return (
                    <label 
                      key={clsName}
                      className={`flex items-center gap-2 p-2.5 rounded-lg border cursor-pointer transition-all ${
                        isChecked 
                          ? 'border-rose-600 bg-rose-50 text-rose-900 font-extrabold' 
                          : 'border-slate-200 bg-white hover:bg-slate-100 text-slate-600'
                      }`}
                    >
                      <input 
                        type="checkbox"
                        checked={isChecked}
                        onChange={() => handleToggleClearClass(clsName)}
                        className="sr-only"
                      />
                      <div 
                        className={`w-4 h-4 rounded border flex items-center justify-center transition-all shrink-0 ${
                          isChecked 
                            ? 'bg-rose-600 border-rose-600 text-white' 
                            : 'border-slate-300 bg-white'
                        }`}
                      >
                        {isChecked && <Check className="w-3 h-3 stroke-[3]" />}
                      </div>
                      <span className="text-xs uppercase tracking-wider">Class {clsName}</span>
                      <span className="text-[10px] font-bold opacity-70 ml-auto">({count})</span>
                    </label>
                  );
                })}
              </div>
            </div>
            
            {/* Action Buttons */}
            <div className="flex flex-col gap-2 pt-2 border-t border-slate-100">
              <button
                onClick={handleClearClassesData}
                disabled={clearSelectedClasses.length === 0}
                className="w-full flex items-center justify-center gap-2 py-3 px-4 rounded-xl font-extrabold text-xs bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-50 disabled:cursor-not-allowed active:scale-[0.98] transition-all shadow-md"
              >
                <Trash2 className="w-4 h-4" />
                Reset Selected Classes ({clearSelectedClasses.length})
              </button>
              
              <button
                onClick={() => setShowClearModal(false)}
                className="w-full py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-600 hover:text-slate-800 text-xs font-bold transition-all"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Create Mentee Room Modal */}
      {showCreateRoomModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-[360px] bg-white border border-slate-200 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 animate-[modalSlideUp_0.2s_ease-out]">
            <div className="flex justify-between items-center pb-2 border-b border-slate-100">
              <h3 className="text-base font-extrabold text-[#333333] flex items-center gap-2">
                <School className="w-5 h-5 text-[#1A365D]" />
                Create Mentee Room
              </h3>
              <button 
                type="button"
                onClick={() => setShowCreateRoomModal(false)}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleCreateRoom} className="flex flex-col gap-3">
              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Room Name
                </label>
                <input
                  type="text"
                  required
                  placeholder="e.g. Advanced Math Guidance Room"
                  value={newRoomName}
                  onChange={(e) => setNewRoomName(e.target.value)}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all font-semibold"
                />
              </div>

              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Mentor / Counselor
                </label>
                <select
                  value={newRoomMentor}
                  onChange={(e) => setNewRoomMentor(e.target.value)}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all font-semibold"
                >
                  {mentors.map(m => (
                    <option key={m.id} value={m.name}>{m.name} ({m.role})</option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                    Target Class
                  </label>
                  <select
                    value={newRoomClass}
                    onChange={(e) => setNewRoomClass(e.target.value)}
                    className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 uppercase focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all font-semibold"
                  >
                    {visibleClasses.map(cls => (
                      <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Description / Goals
                </label>
                <textarea
                  rows="2"
                  placeholder="e.g. Focus on academic support & exam prep"
                  value={newRoomDesc}
                  onChange={(e) => setNewRoomDesc(e.target.value)}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all font-medium resize-none"
                ></textarea>
              </div>

              <div className="flex flex-col gap-2 pt-2 border-t border-slate-100">
                <button
                  type="submit"
                  className="w-full flex items-center justify-center gap-2 py-3 px-4 rounded-xl font-extrabold text-xs bg-[#1A365D] hover:bg-[#2A4365] text-white active:scale-[0.98] transition-all shadow-md"
                >
                  <Plus className="w-4 h-4" />
                  Create Mentee Room
                </button>
                <button
                  type="button"
                  onClick={() => setShowCreateRoomModal(false)}
                  className="w-full py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-600 hover:text-slate-800 text-xs font-bold transition-all"
                >
                  Cancel
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Add Student to Mentee Room Modal */}
      {showAddStudentToRoomModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-[360px] bg-white border border-slate-200 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 animate-[modalSlideUp_0.2s_ease-out]">
            <div className="flex justify-between items-center pb-2 border-b border-slate-100">
              <h3 className="text-base font-extrabold text-[#333333] flex items-center gap-2">
                <UserCheck className="w-5 h-5 text-[#1A365D]" />
                Add Student to Room
              </h3>
              <button 
                type="button"
                onClick={() => setShowAddStudentToRoomModal(false)}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleAddStudentToRoom} className="flex flex-col gap-3">
              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Select Student from Directory
                </label>
                <select
                  required
                  value={selectedStudentForRoom}
                  onChange={(e) => setSelectedStudentForRoom(e.target.value)}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-all font-semibold"
                >
                  <option value="">-- Choose Student --</option>
                  {students.map(s => (
                    <option key={s.id} value={s.id}>
                      {s.name} (Class {s.class.toUpperCase()} - ⭐{s.star})
                    </option>
                  ))}
                </select>
              </div>

              <div className="flex flex-col gap-2 pt-2 border-t border-slate-100">
                <button
                  type="submit"
                  disabled={!selectedStudentForRoom}
                  className="w-full flex items-center justify-center gap-2 py-3 px-4 rounded-xl font-extrabold text-xs bg-[#1A365D] hover:bg-[#2A4365] text-white disabled:opacity-50 disabled:cursor-not-allowed active:scale-[0.98] transition-all shadow-md"
                >
                  <Plus className="w-4 h-4" />
                  Assign to Room
                </button>
                <button
                  type="button"
                  onClick={() => setShowAddStudentToRoomModal(false)}
                  className="w-full py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-600 hover:text-slate-800 text-xs font-bold transition-all"
                >
                  Cancel
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* WhatsApp Report & Share Modal */}
      {showWhatsAppModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-[380px] bg-white border border-slate-200 rounded-2xl p-6 shadow-2xl flex flex-col gap-4 animate-[modalSlideUp_0.2s_ease-out]">
            <div className="flex justify-between items-center pb-2 border-b border-slate-100">
              <h3 className="text-base font-extrabold text-[#333333] flex items-center gap-2">
                <Share2 className="w-5 h-5 text-emerald-600" />
                WhatsApp Report
              </h3>
              <button 
                type="button"
                onClick={() => setShowWhatsAppModal(false)}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div>
              <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                Formatted Report Preview:
              </label>
              <textarea
                readOnly
                rows="8"
                value={whatsAppReportText}
                className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono text-slate-800 focus:outline-none resize-none leading-relaxed"
              ></textarea>
            </div>

            <div className="flex flex-col gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => {
                  const encoded = encodeURIComponent(whatsAppReportText);
                  window.open(`https://api.whatsapp.com/send?text=${encoded}`, '_blank');
                }}
                className="w-full flex items-center justify-center gap-2 py-3 px-4 rounded-xl font-extrabold text-xs bg-emerald-600 hover:bg-emerald-700 text-white active:scale-[0.98] transition-all shadow-md"
              >
                <Share2 className="w-4 h-4" />
                Send Direct via WhatsApp
              </button>

              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(whatsAppReportText);
                  setCopyFeedback('Report copied to clipboard!');
                  setTimeout(() => setCopyFeedback(''), 3000);
                }}
                className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100 text-xs font-bold transition-all"
              >
                <Copy className="w-4 h-4" />
                {copyFeedback || 'Copy Text Only'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Add Mentor Modal */}
      {showAddMentorModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl border border-slate-200 overflow-hidden" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-2 text-[#1A365D]">
                <UserCheck className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider">Add New Mentor</h2>
              </div>
              <button 
                onClick={() => setShowAddMentorModal(false)}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handleAddMentor} className="p-5 flex flex-col gap-4">
              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Mentor Name
                </label>
                <input
                  type="text"
                  required
                  value={newMentor.name}
                  onChange={(e) => setNewMentor({...newMentor, name: e.target.value})}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:border-[#1A365D] focus:bg-white transition-colors"
                  placeholder="e.g. Prof. Jane Doe"
                />
              </div>

              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Room Number
                </label>
                <input
                  type="text"
                  required
                  value={newMentor.roomNumber || ''}
                  onChange={(e) => setNewMentor({...newMentor, roomNumber: e.target.value})}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:border-[#1A365D] focus:bg-white transition-colors"
                  placeholder="e.g. ROOM4"
                />
              </div>

              <div className="pt-2">
                <button
                  type="submit"
                  className="w-full py-3 bg-[#1A365D] hover:bg-[#2A4365] text-white rounded-xl font-extrabold text-xs shadow-md transition-colors active:scale-[0.98]"
                >
                  Create Mentor
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Delete Mentor Modal */}
      {showDeleteMentorModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl border border-slate-200 overflow-hidden" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-2 text-rose-600">
                <Trash2 className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider">Remove Mentor</h2>
              </div>
              <button 
                onClick={() => setShowDeleteMentorModal(false)}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handleDeleteMentor} className="p-5 flex flex-col gap-4">
              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Select Mentor to Remove
                </label>
                <div className="flex flex-col gap-2 max-h-[200px] overflow-y-auto pr-1">
                  {mentors.length > 0 ? mentors.map(m => (
                    <button
                      key={m.id}
                      type="button"
                      onClick={() => setSelectedMentorForDelete(m.id)}
                      className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${
                        selectedMentorForDelete === m.id 
                          ? 'border-rose-500 bg-rose-500/5 ring-1 ring-rose-500/20 shadow-sm' 
                          : 'border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50'
                      }`}
                    >
                      <span className={`text-xs font-bold ${selectedMentorForDelete === m.id ? 'text-rose-700' : 'text-slate-700'}`}>
                        {m.name}
                      </span>
                      {selectedMentorForDelete === m.id && (
                        <CheckCircle2 className="w-4 h-4 text-rose-500" />
                      )}
                    </button>
                  )) : (
                    <p className="text-center text-slate-500 text-xs py-4">No mentors to remove.</p>
                  )}
                </div>
              </div>

              <div className="pt-2">
                <button
                  type="submit"
                  disabled={!selectedMentorForDelete}
                  className="w-full py-3 bg-rose-600 hover:bg-rose-700 disabled:opacity-50 disabled:bg-rose-400 text-white rounded-xl font-extrabold text-xs shadow-md transition-colors active:scale-[0.98]"
                >
                  Delete Selected
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Performance Submit Modal */}
      {showPerformanceSubmitModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl border border-slate-200 overflow-hidden" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-2 text-[#1A365D]">
                <Award className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider uppercase">Submit Score</h2>
              </div>
              <button 
                onClick={() => setShowPerformanceSubmitModal(false)}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handlePerformanceSubmit} className="p-5 flex flex-col gap-4">

              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Count
                </label>
                <input
                  type="number"
                  required
                  min="1"
                  value={performanceSubmitData.count}
                  onChange={(e) => setPerformanceSubmitData({...performanceSubmitData, count: e.target.value})}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:border-[#1A365D] focus:bg-white transition-colors"
                  placeholder="1"
                />
              </div>
              <div className="relative">
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                  Reason
                </label>
                <input
                  type="text"
                  required
                  value={performanceSubmitData.reason}
                  onChange={(e) => {
                    setPerformanceSubmitData({...performanceSubmitData, reason: e.target.value});
                    setShowPerformanceReasonDropdown(true);
                  }}
                  onFocus={() => setShowPerformanceReasonDropdown(true)}
                  onBlur={() => setTimeout(() => setShowPerformanceReasonDropdown(false), 200)}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:border-[#1A365D] focus:bg-white transition-colors"
                  placeholder={`e.g. ${performanceView === 'neat' ? 'Uniform not neat' : performanceView === 'room' ? 'Room untidy' : performanceView === 'spot' ? 'Talking in class' : 'Ineligible behaviour'}`}
                />
                {showPerformanceReasonDropdown && performanceSubmitData.reason && performanceView === 'neat' && (
                  (() => {
                    const filtered = PREDEFINED_NEAT_REASONS.filter(f => f.reason.toLowerCase().includes(performanceSubmitData.reason.toLowerCase()));
                    if (filtered.length === 0) return null;
                    return (
                      <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-slate-200 rounded-xl shadow-lg max-h-48 overflow-y-auto z-50 divide-y divide-slate-100">
                        {filtered.map((item, idx) => (
                          <button
                            key={idx}
                            type="button"
                            onMouseDown={() => {
                              setPerformanceSubmitData({...performanceSubmitData, reason: item.reason, count: item.count});
                              setShowPerformanceReasonDropdown(false);
                            }}
                            className="w-full text-left px-4 py-3 text-sm text-slate-700 hover:bg-[#1A365D]/5 hover:text-[#1A365D] transition-all flex justify-between items-center"
                          >
                            <span className="font-semibold">{item.reason}</span>
                            <span className="text-xs font-extrabold px-2 py-1 rounded-full bg-rose-100 border border-rose-200 text-rose-800">
                              {item.count}
                            </span>
                          </button>
                        ))}
                      </div>
                    );
                  })()
                )}
              </div>
              
              <div className="pt-2">
                <button
                  type="submit"
                  className={`w-full py-3 text-white rounded-xl font-extrabold text-xs shadow-md transition-colors active:scale-[0.98] ${
                    performanceSubmitData.type === 'star' ? 'bg-amber-500 hover:bg-amber-600' : 'bg-rose-500 hover:bg-rose-600'
                  }`}
                >
                  Confirm & Apply to {performanceSelectedStudents.length} Students
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Admin Settings Modal */}
      {showAdminSettingsModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50 shrink-0">
              <div className="flex items-center gap-2 text-[#1A365D]">
                <Settings className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider uppercase">Admin Settings</h2>
              </div>
              <button 
                onClick={() => setShowAdminSettingsModal(false)}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            
            <div className="flex-1 overflow-y-auto p-5 flex flex-col gap-6">
              
              {/* Security / User Management */}
              <div className="flex flex-col gap-3">
                <h3 className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider flex items-center gap-1.5 pb-2 border-b border-slate-100">
                  <Lock className="w-3.5 h-3.5" /> Security
                </h3>
                <div className="text-xs text-slate-500 bg-slate-50 rounded-xl p-3 border border-slate-200">
                  Manage users and passwords via the <strong>Users</strong> tab in the Admin panel.
                </div>
                <button
                  onClick={() => { setShowAdminSettingsModal(false); setAdminSubTab('users'); }}
                  className="px-3 py-2 bg-[#1A365D] text-white text-xs font-bold rounded-xl hover:bg-[#2A4365] transition-colors shadow-sm"
                >
                  Go to User Management →
                </button>
              </div>

              {/* Class Management */}
              <div className="flex flex-col gap-3">
                <h3 className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider flex items-center gap-1.5 pb-2 border-b border-slate-100">
                  <School className="w-3.5 h-3.5" /> Class Management
                </h3>
                <div className="flex justify-between items-center bg-slate-50 p-2 rounded-xl border border-slate-200">
                  <span className="text-[10px] text-slate-500 font-bold px-2">Order & Visibility</span>
                  <div className="flex gap-2">
                    <button onClick={handleAddClass} className="px-2 py-1 bg-emerald-600 text-white text-[10px] font-bold rounded-lg shadow-sm">Add</button>
                  </div>
                </div>
                <div className="flex flex-col gap-2 max-h-60 overflow-y-auto pr-1">
                  {CLASSES.map((cls, idx) => {
                    const isHidden = hiddenClasses.includes(cls);
                    return (
                      <div key={cls} className={`flex items-center justify-between p-2 border rounded-xl transition-all ${isHidden ? 'bg-slate-50 border-slate-200 opacity-60' : 'bg-white border-slate-200 shadow-sm'}`}>
                        <span className={`text-xs font-extrabold uppercase ml-2 ${isHidden ? 'text-slate-400 line-through' : 'text-[#1A365D]'}`}>Class {cls}</span>
                        <div className="flex items-center gap-1">
                          <button
                            onClick={() => handleEditClass(cls)}
                            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100 transition-colors"
                            title="Edit Class Name"
                          >
                            <Pencil className="w-4 h-4" />
                          </button>
                          <button 
                            onClick={() => toggleClassVisibility(cls)}
                            className={`p-1.5 rounded-lg transition-colors ${isHidden ? 'text-slate-400 hover:bg-slate-200' : 'text-emerald-600 hover:bg-emerald-50'}`}
                            title={isHidden ? "Show Class" : "Hide Class"}
                          >
                            {isHidden ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                          </button>
                          <button 
                            onClick={() => setClassToRemoveConfirm(cls)}
                            className="p-1.5 rounded-lg text-rose-500 hover:bg-rose-50 transition-colors"
                            title="Remove Class"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                          <div className="w-px h-4 bg-slate-200 mx-1"></div>
                          <button 
                            onClick={() => moveClassUp(idx)}
                            disabled={idx === 0}
                            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100 disabled:opacity-30 disabled:hover:bg-transparent transition-colors"
                          >
                            <ArrowUp className="w-4 h-4" />
                          </button>
                          <button 
                            onClick={() => moveClassDown(idx)}
                            disabled={idx === CLASSES.length - 1}
                            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100 disabled:opacity-30 disabled:hover:bg-transparent transition-colors"
                          >
                            <ArrowDown className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>


            </div>
            
          </div>
        </div>
      )}



      {/* Add Ineligible Modal */}
      {showAddIneligibleModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50 shrink-0">
              <div className="flex items-center gap-2 text-rose-600">
                <X className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider uppercase">Add Ineligible</h2>
              </div>
              <button 
                onClick={() => setShowAddIneligibleModal(false)}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            
            <form onSubmit={handleAddIneligibleSubmit} className="flex flex-col flex-1 overflow-hidden">
              <div className="p-4 border-b border-slate-200 shrink-0">
                <input
                  type="text"
                  placeholder="Optional Reason (e.g. absent)"
                  value={ineligibleReasonInput}
                  onChange={(e) => setIneligibleReasonInput(e.target.value)}
                  className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-rose-500 transition-colors"
                />
              </div>

              <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-2">
                {students.map(s => {
                  if (s.ineligible) return null; // hide already ineligible
                  const isSelected = ineligibleSelectedStudents.includes(s.id);
                  return (
                    <button 
                      key={s.id}
                      type="button"
                      onClick={() => {
                        setIneligibleSelectedStudents(prev => 
                          prev.includes(s.id) ? prev.filter(id => id !== s.id) : [...prev, s.id]
                        );
                      }}
                      className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${
                        isSelected 
                          ? 'border-rose-600 bg-rose-50 ring-1 ring-rose-600/20 shadow-sm' 
                          : 'border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50'
                      }`}
                    >
                      <div className="flex items-center gap-3">
                        <div className={`w-9 h-9 rounded-full flex items-center justify-center text-[11px] font-extrabold transition-colors ${
                          isSelected ? 'bg-rose-600 text-white' : 'bg-slate-100 text-[#1A365D]'
                        }`}>
                          {getInitials(s.name)}
                        </div>
                        <div>
                          <p className={`text-xs font-extrabold ${isSelected ? 'text-rose-700 uppercase tracking-widest' : 'text-[#333333]'}`}>
                            {s.name}
                          </p>
                          <p className={`text-[10px] font-medium ${isSelected ? 'text-rose-600/70' : 'text-slate-500'}`}>
                            Class {s.class.toUpperCase()}
                          </p>
                        </div>
                      </div>
                      {isSelected && (
                        <div className="w-5 h-5 rounded-full bg-rose-600 text-white flex items-center justify-center shrink-0">
                          <CheckCircle2 className="w-3.5 h-3.5" />
                        </div>
                      )}
                    </button>
                  );
                })}
              </div>
              
              <div className="p-4 border-t border-slate-200 bg-slate-50 shrink-0">
                <button
                  type="submit"
                  disabled={ineligibleSelectedStudents.length === 0}
                  className="w-full py-3.5 bg-rose-600 hover:bg-rose-700 disabled:opacity-50 disabled:bg-rose-400 text-white rounded-xl font-extrabold text-sm shadow-md transition-colors active:scale-[0.98]"
                >
                  Mark {ineligibleSelectedStudents.length} Students Ineligible
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* WhatsApp Summary Modal */}
      {showWhatsappModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl border border-slate-200 overflow-hidden" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-2 text-[#1A365D]">
                <MessageSquare className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider uppercase">Send Summary</h2>
              </div>
              <button 
                onClick={() => {
                  setShowWhatsappModal(false);
                  setPerformanceSelectedClass(null);
                  setPerformanceSelectedStudents([]);
                  setPerformanceView(null);
                  setPerformanceSubmitData({ count: 1, reason: '', type: 'tally' });
                }}
                className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-5 flex flex-col gap-4">
              <p className="text-xs text-slate-500 font-medium leading-relaxed">
                Review the tally summary before downloading or sending via WhatsApp:
              </p>
              <div id="whatsapp-summary-card" className="bg-slate-50 p-4 rounded-xl border border-slate-200 text-xs font-mono text-slate-700 whitespace-pre-wrap h-48 overflow-y-auto shadow-inner">
                {whatsappMessage}
              </div>
              
              <div className="pt-2 flex flex-col sm:flex-row gap-2">
                <button
                  onClick={() => handleDownloadScreenshotCard('whatsapp-summary-card', 'Summary_Report.png')}
                  className="flex-1 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-extrabold text-xs shadow-md transition-colors active:scale-[0.98] flex justify-center items-center gap-2 min-h-[44px]"
                >
                  <Camera className="w-4 h-4" />
                  Screenshot
                </button>
                <a
                  href={`https://wa.me/?text=${encodeURIComponent(whatsappMessage)}`}
                  target="_blank"
                  rel="noreferrer"
                  onClick={() => {
                    setShowWhatsappModal(false);
                    setPerformanceSelectedClass(null);
                    setPerformanceSelectedStudents([]);
                    setPerformanceView(null);
                    setPerformanceSubmitData({ count: 1, reason: '', type: 'tally' });
                  }}
                  className="flex-1 py-3 bg-[#25D366] hover:bg-[#1DA851] text-white rounded-xl font-extrabold text-xs shadow-md transition-colors active:scale-[0.98] flex justify-center items-center gap-2 min-h-[44px]"
                >
                  <MessageSquare className="w-4 h-4" />
                  WhatsApp
                </a>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* IR Date Selection Modal */}
      {showIRDateModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[105] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl border border-slate-200 overflow-hidden flex flex-col scale-in">
            <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between">
              <div className="flex items-center gap-2 text-[#1A365D]">
                <Calendar className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider uppercase">Select Date Range</h2>
              </div>
              <button onClick={() => setShowIRDateModal(false)} className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors">
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-5 flex-1 bg-slate-50/50 flex flex-col gap-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-2">From Date</label>
                <input
                  type="date"
                  value={irFromDate}
                  onChange={(e) => setIrFromDate(e.target.value)}
                  className="w-full px-4 py-3 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all font-medium text-slate-700"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-2">To Date</label>
                <input
                  type="date"
                  value={irToDate}
                  onChange={(e) => setIrToDate(e.target.value)}
                  className="w-full px-4 py-3 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all font-medium text-slate-700"
                />
              </div>
              <button
                onClick={() => {
                  setShowIRDateModal(false);
                  if (dateModalNextAction === 'ADMIN_REPORT') {
                    setShowDownloadModal(true);
                  } else {
                    setShowIRSelectModal(true);
                  }
                }}
                className="w-full mt-2 bg-[#1A365D] hover:bg-blue-800 text-white font-bold py-3 rounded-xl shadow-lg shadow-blue-900/20 transition-all flex items-center justify-center gap-2 uppercase tracking-wide text-sm"
              >
                <span>OK, Continue</span>
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* IR Student Selection Modal */}
      {showIRSelectModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[80vh]" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50 shrink-0">
              <div className="flex items-center gap-2 text-[#1A365D]">
                <FileText className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider uppercase">Select Student for IR</h2>
              </div>
              <button onClick={() => setShowIRSelectModal(false)} className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors">
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-4 overflow-y-auto flex-1 bg-slate-50/50">
              <div className="grid grid-cols-1 gap-2">
                <button
                  onClick={() => {
                    setSelectedIRStudent('ALL');
                    setShowIRSelectModal(false);
                    setShowIRModal(true);
                  }}
                  className="w-full text-left px-4 py-3 bg-white hover:bg-blue-50 border border-slate-200 rounded-xl font-bold text-[#1A365D] transition-colors shadow-sm flex justify-between items-center"
                >
                  <span>Select All Students</span>
                  <ChevronRight className="w-4 h-4 opacity-50" />
                </button>
                {irAssignedStudents.map(student => (
                  <button
                    key={student.id}
                    onClick={() => {
                      setSelectedIRStudent(student);
                      setShowIRSelectModal(false);
                      setShowIRModal(true);
                    }}
                    className="w-full text-left px-4 py-3 bg-white hover:bg-slate-50 border border-slate-200 rounded-xl flex justify-between items-center transition-colors shadow-sm"
                  >
                    <div>
                      <div className="font-bold text-slate-800">{student.name}</div>
                      <div className="text-xs text-slate-500 font-medium">Class {student.class}</div>
                    </div>
                    <ChevronRight className="w-4 h-4 text-slate-400" />
                  </button>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Individual Report Modal */}
      {showIRModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[110] animate-fade-in">
          <div className="bg-slate-100 rounded-2xl w-full max-w-4xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col h-[90vh]" onClick={e => e.stopPropagation()}>
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-white shrink-0">
              <div className="flex items-center gap-2 text-[#1A365D]">
                <FileText className="w-5 h-5" />
                <h2 className="text-sm font-extrabold tracking-wider uppercase">Individual Report</h2>
              </div>
              <div className="flex gap-2">
                <button
                  onClick={() => {
                    // Excel download logic for IR
                    const csvLines = ["Date,Name,Class,Event Item,Reason / Notes"];
                    
                    const studentsToExport = selectedIRStudent === 'ALL' ? irAssignedStudents : [selectedIRStudent];
                    
                    studentsToExport.forEach(s => {
                      const studentLogs = irHistoryLogs
                        .filter(log => log.student_id === s.id)
                        .filter(log => {
                          const logDate = new Date(log.date).toISOString().split('T')[0];
                          return (!irFromDate || logDate >= irFromDate) && (!irToDate || logDate <= irToDate);
                        });
                      if (studentLogs.length > 0) {
                        studentLogs.forEach(log => {
                          const d = new Date(log.date);
                          const formattedDate = d.toLocaleDateString('en-GB');
                          let eventDisplay = '';
                          const type = log.event_type.toLowerCase();
                          if (type.includes('sheet') || type === 'apology') {
                            eventDisplay = log.event_type;
                          } else if (type === 'n&o' || type === 'n&o tally' || type === 'spot fine') {
                            eventDisplay = log.event_type;
                          } else {
                            eventDisplay = `${Math.abs(log.amount)} ${log.event_type}`;
                          }
                          const reasonSafe = (log.reason || '').replace(/"/g, '""').replace(/\n/g, ' ');
                          csvLines.push(`${formattedDate},"${s.name.replace(/"/g, '""')}",${s.class},"${eventDisplay}","${reasonSafe}"`);
                        });
                      } else {
                        const today = new Date().toLocaleDateString('en-GB');
                        csvLines.push(`${today},"${s.name.replace(/"/g, '""')}",${s.class},"No events","No history records found"`);
                      }
                    });
                    
                    const csvContent = csvLines.join('\n');
                    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement("a");
                    link.setAttribute("href", url);
                    link.setAttribute("download", `Individual_Report_${selectedIRStudent === 'ALL' ? 'All' : selectedIRStudent.name.replace(/\\s+/g, '_')}.csv`);
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                  }}
                  className="px-3 py-1.5 bg-emerald-500 hover:bg-emerald-600 text-white text-[10px] uppercase font-bold rounded-lg shadow-sm transition-colors flex items-center gap-1"
                >
                  <span className="text-sm">📥</span> Download Excel
                </button>
                <button onClick={() => setShowIRModal(false)} className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors">
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>
            
            <div className="p-6 overflow-y-auto flex-1 flex flex-col gap-8">
              {(selectedIRStudent === 'ALL' ? irAssignedStudents : [selectedIRStudent]).map(student => (
                <div key={student.id} className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col gap-6">
                  {/* Header: Image, Name, Class */}
                  <div className="flex flex-col md:flex-row items-start md:items-center gap-6">
                    <div className="w-24 h-24 bg-slate-800 rounded-xl flex items-center justify-center shrink-0 shadow-inner">
                      <UserCheck className="w-10 h-10 text-slate-500 opacity-50" />
                    </div>
                    <div className="flex flex-col">
                      <h1 className="text-3xl font-black text-[#1A365D] uppercase tracking-tight">{student.name}</h1>
                      <div className="flex items-center gap-2 mt-2">
                        <span className="px-3 py-1 bg-blue-100 text-blue-800 rounded-lg text-sm font-bold uppercase tracking-wider">
                          Class {student.class}
                        </span>
                      </div>
                    </div>
                  </div>
                  
                  {/* Totals Section */}
                  <div className="grid grid-cols-2 md:grid-cols-6 gap-3 mt-2">
                    <div className="bg-emerald-50 border border-emerald-100 rounded-xl p-4 flex flex-col items-center justify-center text-center">
                      <div className="text-xs font-bold text-emerald-600 uppercase tracking-wider mb-1">Total Star</div>
                      <div className="text-2xl font-black text-emerald-700">{student.star || 0}</div>
                    </div>
                    <div className="bg-rose-50 border border-rose-100 rounded-xl p-4 flex flex-col items-center justify-center text-center">
                      <div className="text-xs font-bold text-rose-600 uppercase tracking-wider mb-1">Total Tally</div>
                      <div className="text-2xl font-black text-rose-700">{student.tally || 0}</div>
                    </div>
                    <div className="bg-red-50 border border-red-100 rounded-xl p-4 flex flex-col items-center justify-center text-center">
                      <div className="text-xs font-bold text-red-600 uppercase tracking-wider mb-1">Total Fine</div>
                      <div className="text-2xl font-black text-red-700">₹{student.fine || 0}</div>
                    </div>
                    <div className="bg-amber-50 border border-amber-100 rounded-xl p-4 flex flex-col items-center justify-center text-center">
                      <div className="text-xs font-bold text-amber-600 uppercase tracking-wider mb-1">N&O Tally</div>
                      <div className="text-2xl font-black text-amber-700">{student.neatAndOrderTally || 0}</div>
                    </div>
                    <div className="bg-purple-50 border border-purple-100 rounded-xl p-4 flex flex-col items-center justify-center text-center">
                      <div className="text-xs font-bold text-purple-600 uppercase tracking-wider mb-1">Diary Tally</div>
                      <div className="text-2xl font-black text-purple-700">{student.diaryTally || 0}</div>
                    </div>
                    <div className="bg-cyan-50 border border-cyan-100 rounded-xl p-4 flex flex-col items-center justify-center text-center">
                      <div className="text-xs font-bold text-cyan-600 uppercase tracking-wider mb-1">Sheet Tally</div>
                      <div className="text-2xl font-black text-cyan-700">{student.sheetTally || 0}</div>
                    </div>
                  </div>
                  
                  {/* History Table */}
                  <div className="mt-2">
                    <h3 className="text-sm font-extrabold text-slate-800 mb-3 uppercase tracking-wider flex items-center gap-2">
                      <BarChart3 className="w-4 h-4 text-slate-400" />
                      Report History
                    </h3>
                    <div className="overflow-x-auto rounded-xl border border-slate-200">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-100">
                            <th className="p-3 text-xs font-bold text-slate-600 uppercase tracking-wider border-b border-slate-200 w-1/4">Date</th>
                            <th className="p-3 text-xs font-bold text-slate-600 uppercase tracking-wider border-b border-slate-200 w-1/4">Event Item</th>
                            <th className="p-3 text-xs font-bold text-slate-600 uppercase tracking-wider border-b border-slate-200 w-1/2">Reason / Notes</th>
                          </tr>
                        </thead>
                        <tbody className="bg-white">
                          {irHistoryLogs
                            .filter(log => log.student_id === student.id)
                            .filter(log => {
                              const logDate = new Date(log.date).toISOString().split('T')[0];
                              return (!irFromDate || logDate >= irFromDate) && (!irToDate || logDate <= irToDate);
                            }).length > 0 ? (
                            irHistoryLogs
                              .filter(log => log.student_id === student.id)
                              .filter(log => {
                                const logDate = new Date(log.date).toISOString().split('T')[0];
                                return (!irFromDate || logDate >= irFromDate) && (!irToDate || logDate <= irToDate);
                              })
                              .map(log => {
                              const d = new Date(log.date);
                              const formattedDate = d.toLocaleDateString('en-GB');
                              
                              let eventDisplay = '';
                              const type = log.event_type.toLowerCase();
                              if (type.includes('sheet') || type === 'apology') {
                                eventDisplay = log.event_type;
                              } else if (type === 'n&o' || type === 'n&o tally' || type === 'spot fine') {
                                eventDisplay = log.event_type;
                              } else {
                                eventDisplay = `${Math.abs(log.amount)} ${log.event_type}`;
                              }
                              
                              return (
                                <tr key={log.id} className="hover:bg-slate-50 transition-colors">
                                  <td className="p-3 text-sm font-medium text-slate-700 border-b border-slate-100">{formattedDate}</td>
                                  <td className="p-3 text-sm font-bold text-[#1A365D] border-b border-slate-100 capitalize">{eventDisplay}</td>
                                  <td className="p-3 text-sm text-slate-600 border-b border-slate-100">{log.reason || '-'}</td>
                                </tr>
                              );
                            })
                          ) : (
                            <tr>
                              <td colSpan="3" className="p-6 text-center text-slate-400 font-medium italic">
                                No history records found for this student.
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
      {classToRemoveConfirm && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[200] animate-fade-in">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm overflow-hidden flex flex-col scale-in">
            <div className="p-4 border-b border-slate-100 flex items-center gap-3">
              <div className="w-10 h-10 rounded-full bg-rose-100 flex items-center justify-center shrink-0">
                <AlertCircle className="w-5 h-5 text-rose-600" />
              </div>
              <div>
                <h3 className="font-black tracking-tight text-slate-800 text-lg">Warning!</h3>
                <p className="text-xs font-bold text-slate-500 uppercase tracking-wider">Remove Class & Students</p>
              </div>
            </div>
            <div className="p-5 text-sm text-slate-600">
              Are you sure you want to remove <strong>CLASS {classToRemoveConfirm.toUpperCase()}</strong>?
              <br /><br />
              <span className="text-rose-600 font-bold">This will DELETE the class AND ALL students in this class. All their data will be permanently deleted.</span>
            </div>
            <div className="p-4 bg-slate-50 flex gap-3 justify-end border-t border-slate-100">
              <button
                onClick={() => setClassToRemoveConfirm(null)}
                className="px-4 py-2 text-slate-600 font-bold text-sm bg-white border border-slate-200 rounded-xl hover:bg-slate-50 active:scale-95 transition-all"
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  setCLASSES(CLASSES.filter(c => c !== classToRemoveConfirm));
                  setStudents(students.filter(s => s.class !== classToRemoveConfirm));
                  fetch(`/api/classes/${classToRemoveConfirm}`, { method: 'DELETE' }).catch(console.error);
                  setClassToRemoveConfirm(null);
                }}
                className="px-4 py-2 bg-rose-600 text-white font-bold text-sm rounded-xl hover:bg-rose-700 shadow-sm shadow-rose-200 active:scale-95 transition-all"
              >
                OK
              </button>
            </div>
          </div>
        </div>
      )}
      {/* ROOM TALLY MODAL */}
      {showRoomTallyModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[120] animate-fade-in">
          <div className="bg-white rounded-3xl w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-6 gap-5" onClick={e => e.stopPropagation()}>
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2 text-emerald-700">
                <Award className="w-6 h-6" />
                <h3 className="font-extrabold text-lg">Room Tally Entry</h3>
              </div>
              <button onClick={() => setShowRoomTallyModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
            </div>

            <div className="flex flex-col gap-4">
              <div>
                <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">Tally Number (Count)</label>
                <div className="flex items-center gap-3">
                  <button
                    onClick={() => setRoomTallyCount(prev => Math.max(1, prev - 1))}
                    className="w-11 h-11 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-black text-lg flex items-center justify-center transition-colors"
                  >
                    -
                  </button>
                  <input
                    type="number"
                    min="1"
                    value={roomTallyCount}
                    onChange={e => setRoomTallyCount(Math.max(1, parseInt(e.target.value) || 1))}
                    className="flex-1 py-2.5 px-4 bg-slate-50 border border-slate-200 rounded-xl font-black text-center text-xl text-slate-800 focus:outline-none focus:border-emerald-500"
                  />
                  <button
                    onClick={() => setRoomTallyCount(prev => prev + 1)}
                    className="w-11 h-11 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-black text-lg flex items-center justify-center transition-colors shadow-xs"
                  >
                    +
                  </button>
                </div>
              </div>

              <div>
                <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">Reason / Notes</label>
                <input
                  type="text"
                  value={roomTallyReason}
                  onChange={e => setRoomTallyReason(e.target.value)}
                  placeholder="e.g. Room untidy, Bed not properly made..."
                  className="w-full py-3 px-4 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-sm text-slate-800 focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div className="p-3 bg-emerald-50 border border-emerald-100 rounded-xl text-xs text-emerald-800 font-bold">
                Applying to {performanceSelectedStudents.length} selected student(s) in {selectedRoom}
              </div>
            </div>

            <div className="flex items-center gap-3 pt-2">
              <button
                onClick={() => setShowRoomTallyModal(false)}
                className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  const currentRoomKey = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
                  const allocated = ROOM_STUDENT_MAPPING[currentRoomKey] !== undefined ? 
                    ROOM_STUDENT_MAPPING[currentRoomKey] : 
                    students.filter(s => String(s.room || s.roomNumber || '').trim() === currentRoomKey);
                  const selectedObjs = allocated.filter(s => performanceSelectedStudents.includes(s.id));
                  
                  const amount = Number(roomTallyCount) || 1;
                  const reason = roomTallyReason || 'Room Tally';

                  let studentsToUpsert = [];
                  const updatedStudents = students.map(s => {
                    const isMatchingSelected = selectedObjs.some(so => 
                      so.id === s.id || (so.name.toLowerCase() === s.name.toLowerCase() && so.class.toLowerCase() === (s.class || '').toLowerCase())
                    );
                    if (isMatchingSelected) {
                      const updated = {
                        ...s,
                        tally: (s.tally || 0) + amount,
                        neatAndOrderTally: (s.neatAndOrderTally || 0) + amount,
                        neatAndOrderIncidents: (s.neatAndOrderIncidents || 0) + 1,
                        neatAndOrderReason: reason || s.neatAndOrderReason
                      };
                      logHistory(s.id, 'N&O Tally', amount, reason);
                      studentsToUpsert.push(updated);
                      return updated;
                    }
                    return s;
                  });

                  if (studentsToUpsert.length > 0) {
                    setStudents(updatedStudents);
                    fetch('/api/students/bulk-upsert', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ students: studentsToUpsert })
                    }).catch(err => console.error("Error bulk upserting room tally:", err));
                  }

                  const now = new Date();
                  const summary = {
                    date: now.toLocaleDateString('en-GB'),
                    time: now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                    room: `${selectedHostel} • ${selectedRoom}`,
                    tallyNumber: amount,
                    reason: reason,
                    students: selectedObjs
                  };
                  
                  setShowRoomTallyModal(false);
                  setRoomTallySummary(summary);
                }}
                className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95"
              >
                DONE
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ROOM TALLY SUMMARY MODAL */}
      {roomTallySummary && (
        <div className="fixed inset-0 bg-slate-900/65 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 z-[130] animate-fade-in overflow-y-auto">
          <div className="bg-white rounded-3xl w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-5 sm:p-6 gap-4 sm:gap-5 max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
            <div id="room-tally-summary-card" className="flex flex-col gap-4 bg-white p-3 rounded-2xl border border-slate-100">
              <div className="bg-gradient-to-r from-emerald-600 to-teal-600 p-4 rounded-2xl text-white text-center shadow-md">
                <span className="text-[10px] font-black uppercase tracking-widest text-emerald-200 block">Performance Summary</span>
                <h3 className="text-xl font-black uppercase mt-0.5">ROOM TALLY ENTRY</h3>
              </div>

              <div className="space-y-3 bg-slate-50 p-4 rounded-2xl border border-slate-200/80 text-xs">
                <div className="flex justify-between border-b border-slate-200 pb-2">
                  <span className="font-bold text-slate-500">Date:</span>
                  <span className="font-extrabold text-slate-800">{roomTallySummary.date}</span>
                </div>
                <div className="flex justify-between border-b border-slate-200 pb-2">
                  <span className="font-bold text-slate-500">Time:</span>
                  <span className="font-extrabold text-slate-800">{roomTallySummary.time}</span>
                </div>
                <div className="flex justify-between border-b border-slate-200 pb-2">
                  <span className="font-bold text-slate-500">Room:</span>
                  <span className="font-extrabold text-emerald-700">{roomTallySummary.room}</span>
                </div>
                <div className="flex justify-between border-b border-slate-200 pb-2">
                  <span className="font-bold text-slate-500">Tally Number:</span>
                  <span className="font-black text-emerald-600 text-sm">{roomTallySummary.tallyNumber}</span>
                </div>
                <div className="flex justify-between pb-1">
                  <span className="font-bold text-slate-500">Reason:</span>
                  <span className="font-extrabold text-slate-800 text-right">{roomTallySummary.reason}</span>
                </div>
              </div>

              <div>
                <span className="text-[11px] font-extrabold text-slate-400 uppercase tracking-wider block mb-2">Allocated Students</span>
                <div className="max-h-36 overflow-y-auto space-y-1.5 pr-1">
                  {roomTallySummary.students.map((s, idx) => (
                    <div key={s.id || idx} className="p-2.5 bg-slate-50 border border-slate-200 rounded-xl flex justify-between items-center text-xs">
                      <span className="font-bold text-slate-800">{idx + 1}. {s.name}</span>
                      <span className="font-black text-emerald-700 text-[10px] bg-emerald-100 px-2 py-0.5 rounded-md">Class {s.class}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 sm:gap-3 pt-2">
              <button
                onClick={() => handleDownloadScreenshotCard('room-tally-summary-card', `Room_Tally_${roomTallySummary.room}.png`)}
                className="flex-1 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-2 min-h-[44px]"
              >
                <Camera className="w-4 h-4" />
                <span>Download Screenshot</span>
              </button>
              <button
                onClick={() => {
                  const studentListStr = roomTallySummary.students.map((s, i) => `${i + 1}. ${s.name} (${s.class})`).join('\n');
                  const msg = `*ROOM TALLY ENTRY SUMMARY*\n\n` +
                    `*Date:* ${roomTallySummary.date}\n` +
                    `*Time:* ${roomTallySummary.time}\n` +
                    `*Room:* ${roomTallySummary.room}\n` +
                    `*Tally Number:* ${roomTallySummary.tallyNumber}\n` +
                    `*Reason:* ${roomTallySummary.reason}\n\n` +
                    `*Students Allocated:*\n${studentListStr}`;
                  
                  window.open(`https://api.whatsapp.com/send?text=${encodeURIComponent(msg)}`, '_blank');
                }}
                className="flex-1 py-3 bg-[#25D366] hover:bg-[#1DA851] text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-2 min-h-[44px]"
              >
                <span>Share WhatsApp</span>
              </button>
              <button
                onClick={() => setRoomTallySummary(null)}
                className="py-3 px-5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors min-h-[44px]"
              >
                RETURN
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ROOM FINE MODAL */}
      {showRoomFineModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[120] animate-fade-in">
          <div className="bg-white rounded-3xl w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-6 gap-5" onClick={e => e.stopPropagation()}>
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2 text-amber-600">
                <AlertTriangle className="w-6 h-6" />
                <h3 className="font-extrabold text-lg">Room Fine Entry</h3>
              </div>
              <button onClick={() => setShowRoomFineModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
            </div>

            <div className="flex flex-col gap-4">
              <div>
                <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">Fine Amount (₹)</label>
                <input
                  type="number"
                  value={roomFineAmount}
                  onChange={e => setRoomFineAmount(e.target.value)}
                  placeholder="e.g. 50, 100, 200"
                  className="w-full py-3 px-4 bg-slate-50 border border-slate-200 rounded-xl font-black text-lg text-slate-800 focus:outline-none focus:border-amber-500"
                />
              </div>

              <div>
                <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">Reason / Violation</label>
                <input
                  type="text"
                  value={roomFineReason}
                  onChange={e => setRoomFineReason(e.target.value)}
                  placeholder="e.g. Late night disturbance, Damaged property..."
                  className="w-full py-3 px-4 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-sm text-slate-800 focus:outline-none focus:border-amber-500"
                />
              </div>

              <div className="p-3 bg-amber-50 border border-amber-100 rounded-xl text-xs text-amber-800 font-bold">
                Applying to {performanceSelectedStudents.length} selected student(s) in {selectedRoom}
              </div>
            </div>

            <div className="flex items-center gap-3 pt-2">
              <button
                onClick={() => setShowRoomFineModal(false)}
                className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  const currentRoomKey = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
                  const allocated = ROOM_STUDENT_MAPPING[currentRoomKey] !== undefined ? 
                    ROOM_STUDENT_MAPPING[currentRoomKey] : 
                    students.filter(s => String(s.room || s.roomNumber || '').trim() === currentRoomKey);
                  const selectedObjs = allocated.filter(s => performanceSelectedStudents.includes(s.id));
                  
                  const fineAmt = Number(roomFineAmount) || 0;
                  const reason = roomFineReason || 'Room Fine Violation';

                  let studentsToUpsert = [];
                  const updatedStudents = students.map(s => {
                    const isMatchingSelected = selectedObjs.some(so => 
                      so.id === s.id || (so.name.toLowerCase() === s.name.toLowerCase() && so.class.toLowerCase() === (s.class || '').toLowerCase())
                    );
                    if (isMatchingSelected) {
                      const updated = {
                        ...s,
                        spotFine: (s.spotFine || 0) + fineAmt,
                        spotFineReason: reason || s.spotFineReason
                      };
                      logHistory(s.id, 'Spot Fine', fineAmt, reason);
                      studentsToUpsert.push(updated);
                      return updated;
                    }
                    return s;
                  });

                  if (studentsToUpsert.length > 0) {
                    setStudents(updatedStudents);
                    fetch('/api/students/bulk-upsert', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ students: studentsToUpsert })
                    }).catch(err => console.error("Error bulk upserting room fine:", err));
                  }

                  const now = new Date();
                  const summary = {
                    date: now.toLocaleDateString('en-GB'),
                    time: now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                    room: `${selectedHostel} • ${selectedRoom}`,
                    amount: fineAmt,
                    reason: reason,
                    students: selectedObjs
                  };
                  
                  setShowRoomFineModal(false);
                  setRoomFineSummary(summary);
                }}
                className="flex-1 py-3 bg-amber-500 hover:bg-amber-600 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95"
              >
                DONE
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ROOM FINE SUMMARY MODAL */}
      {roomFineSummary && (
        <div className="fixed inset-0 bg-slate-900/65 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 z-[130] animate-fade-in overflow-y-auto">
          <div className="bg-white rounded-3xl w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-5 sm:p-6 gap-4 sm:gap-5 max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
            <div id="room-fine-summary-card" className="flex flex-col gap-4 bg-white p-3 rounded-2xl border border-slate-100">
              <div className="bg-gradient-to-r from-amber-500 to-orange-600 p-4 rounded-2xl text-white text-center shadow-md">
                <span className="text-[10px] font-black uppercase tracking-widest text-amber-100 block">Performance Summary</span>
                <h3 className="text-xl font-black uppercase mt-0.5">ROOM FINE ENTRY</h3>
              </div>

              <div className="space-y-3 bg-slate-50 p-4 rounded-2xl border border-slate-200/80 text-xs">
                <div className="flex justify-between border-b border-slate-200 pb-2">
                  <span className="font-bold text-slate-500">Date:</span>
                  <span className="font-extrabold text-slate-800">{roomFineSummary.date}</span>
                </div>
                <div className="flex justify-between border-b border-slate-200 pb-2">
                  <span className="font-bold text-slate-500">Time:</span>
                  <span className="font-extrabold text-slate-800">{roomFineSummary.time}</span>
                </div>
                <div className="flex justify-between border-b border-slate-200 pb-2">
                  <span className="font-bold text-slate-500">Room:</span>
                  <span className="font-extrabold text-amber-700">{roomFineSummary.room}</span>
                </div>
                <div className="flex justify-between border-b border-slate-200 pb-2">
                  <span className="font-bold text-slate-500">Fine Amount:</span>
                  <span className="font-black text-amber-600 text-sm">₹{roomFineSummary.amount}</span>
                </div>
                <div className="flex justify-between pb-1">
                  <span className="font-bold text-slate-500">Reason:</span>
                  <span className="font-extrabold text-slate-800 text-right">{roomFineSummary.reason}</span>
                </div>
              </div>

              <div>
                <span className="text-[11px] font-extrabold text-slate-400 uppercase tracking-wider block mb-2">Allocated Students</span>
                <div className="max-h-36 overflow-y-auto space-y-1.5 pr-1">
                  {roomFineSummary.students.map((s, idx) => (
                    <div key={s.id || idx} className="p-2.5 bg-slate-50 border border-slate-200 rounded-xl flex justify-between items-center text-xs">
                      <span className="font-bold text-slate-800">{idx + 1}. {s.name}</span>
                      <span className="font-black text-amber-700 text-[10px] bg-amber-100 px-2 py-0.5 rounded-md">Class {s.class}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 sm:gap-3 pt-2">
              <button
                onClick={() => handleDownloadScreenshotCard('room-fine-summary-card', `Room_Fine_${roomFineSummary.room}.png`)}
                className="flex-1 py-3 bg-amber-600 hover:bg-amber-700 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-2 min-h-[44px]"
              >
                <Camera className="w-4 h-4" />
                <span>Download Screenshot</span>
              </button>
              <button
                onClick={() => {
                  const studentListStr = roomFineSummary.students.map((s, i) => `${i + 1}. ${s.name} (${s.class})`).join('\n');
                  const msg = `*ROOM FINE ENTRY SUMMARY*\n\n` +
                    `*Date:* ${roomFineSummary.date}\n` +
                    `*Time:* ${roomFineSummary.time}\n` +
                    `*Room:* ${roomFineSummary.room}\n` +
                    `*Fine Amount:* ₹${roomFineSummary.amount}\n` +
                    `*Reason:* ${roomFineSummary.reason}\n\n` +
                    `*Students Allocated:*\n${studentListStr}`;
                  
                  window.open(`https://api.whatsapp.com/send?text=${encodeURIComponent(msg)}`, '_blank');
                }}
                className="flex-1 py-3 bg-[#25D366] hover:bg-[#1DA851] text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-2 min-h-[44px]"
              >
                <span>Share WhatsApp</span>
              </button>
              <button
                onClick={() => setRoomFineSummary(null)}
                className="py-3 px-5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors min-h-[44px]"
              >
                BACK
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
