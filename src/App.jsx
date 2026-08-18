import React, { useState, useEffect, useRef } from 'react';
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
  Camera,
  Eye,
  EyeOff,
  ArrowUp,
  ArrowDown,
  Pencil,
  Sun,
  List,
  Smartphone,
  FileSpreadsheet,
  Upload,
  UserPlus,
  Clock,
  PhoneCall
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
    { id: 'c1b-39-1785606211302', name: 'SHAZIN NAUSHAD', class: 'C1B' }
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
    { id: 'r109-3', name: 'AMAN ABDULLA KT', class: 'C1C' },
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
  ],
  // WAVES Hostel Block Rooms Allotments (A02, A03, A04, A05, A06, A08, A09, A10, A11, A13)
  'WAVES_A04': [
    { id: 'w-a04-1', name: 'ABDUL FATHAH FIROS', class: 'JR9' },
    { id: 'w-a04-2', name: 'MUHAMMAD AMEEN IBRAHIM', class: 'JR8' },
    { id: 'w-a04-3', name: 'NITHASH ALI M', class: 'JR8' },
    { id: 'w-a04-4', name: 'NOUMAN ALI', class: 'JR8' },
    { id: 'w-a04-5', name: 'HIFAZ BIN MUHAMMED AFTHAB', class: 'JR9' },
    { id: 'w-a04-6', name: 'FAZAL AHMED M', class: 'JR8' },
    { id: 'w-a04-7', name: 'ABDULLA BIN FAISAL', class: 'JR9' },
    { id: 'w-a04-8', name: 'HAFI YOOSUF A T', class: 'JR9' }
  ],
  'WAVES_A05': [
    { id: 'w-a05-1', name: 'AAMIR BIN USMAN KEELATH', class: 'JR9' },
    { id: 'w-a05-2', name: 'MIDLAJ T P', class: 'JR8' },
    { id: 'w-a05-3', name: 'MISBAH ABDULLAH S', class: 'JR8' },
    { id: 'w-a05-4', name: 'MOHAMMED SWALIH', class: 'JR9' },
    { id: 'w-a05-5', name: 'MUHAMMED ZAYAN C', class: 'JR9' },
    { id: 'w-a05-6', name: 'NAASHITH ALI C K', class: 'JR9' }
  ],
  'WAVES_A06': [
    { id: 'w-a06-1', name: 'MUHAMMED RAZI K', class: 'JR9' },
    { id: 'w-a06-2', name: 'AMAL RAHMAN M.P', class: 'JR8' },
    { id: 'w-a06-3', name: 'AZIN .E', class: 'JR8' },
    { id: 'w-a06-4', name: 'MUHAMMED HISHAM', class: 'JR9' },
    { id: 'w-a06-5', name: 'FEMIL MUHAMMED T', class: 'JR9' },
    { id: 'w-a06-6', name: 'MUHAMMED YAMIN ANSARI M A', class: 'JR9' },
    { id: 'w-a06-7', name: 'MUHAMMED SAJAD', class: 'JR9' },
    { id: 'w-a06-8', name: 'MOHAMMED AZEEM V', class: 'JR8' }
  ],
  'WAVES_A10': [
    { id: 'w-a10-1', name: 'SHAFIN MUHAMMED', class: 'JR8' },
    { id: 'w-a10-2', name: 'TA ALI ZAKWAN', class: 'JR8' },
    { id: 'w-a10-3', name: 'IMAN MUHAMMAD P T', class: 'JR9' },
    { id: 'w-a10-4', name: 'YAZAN MUHAMMED', class: 'JR9' },
    { id: 'w-a10-5', name: 'MUHAMMED MUBARIZ CK', class: 'JR8' },
    { id: 'w-a10-6', name: 'HAFIZ MUHAMMAD K', class: 'JR8' }
  ],
  'WAVES_A02': [
    { id: 'w-a02-1', name: 'NAHYAN ABDULLA', class: 'C2B' },
    { id: 'w-a02-2', name: 'MOHAMMED AMEEN PK', class: 'C2C' },
    { id: 'w-a02-3', name: 'FADL AHMED ZULFIQAR ALI', class: 'C2B' },
    { id: 'w-a02-4', name: 'MUHAMMAD HAMDAN', class: 'C2B' },
    { id: 'w-a02-5', name: 'SHAMMAS ASHRAF', class: 'C2C' },
    { id: 'w-a02-6', name: 'MUHAMMED NAZAL CK', class: 'C2C' },
    { id: 'w-a02-7', name: 'MUHAMMED AMAN VK', class: 'C2C' },
    { id: 'w-a02-8', name: 'MUHAMMED SAFWAN MM', class: 'C2C' }
  ],
  'WAVES_A03': [
    { id: 'w-a03-1', name: 'RAYAAN MUHAMMED', class: 'C2B' },
    { id: 'w-a03-2', name: 'SYED RAAHIL', class: 'C2B' },
    { id: 'w-a03-3', name: 'MUHAMMED SINAN VK', class: 'C2B' },
    { id: 'w-a03-4', name: 'MINHAJ', class: 'C2C' },
    { id: 'w-a03-5', name: 'ABDUNNOOR P', class: 'C2B' },
    { id: 'w-a03-6', name: 'MUHAMMED AMEEN TK', class: 'C2C' },
    { id: 'w-a03-7', name: 'SHAHRAN', class: 'C2B' },
    { id: 'w-a03-8', name: 'UMMAR SHAIKHALI', class: 'C2C' }
  ],
  'WAVES_A08': [
    { id: 'w-a08-1', name: 'MUHAMMED NASH TT', class: 'C2C' },
    { id: 'w-a08-2', name: 'MUAD CK', class: 'C2C' },
    { id: 'w-a08-3', name: 'MAIZ MUHAMMED', class: 'C2C' },
    { id: 'w-a08-4', name: 'MUHAMMED SINAN K CHEKIAD', class: 'C2B' },
    { id: 'w-a08-5', name: 'MOHAMMAD ZABIN SHAMSEER', class: 'C2C' },
    { id: 'w-a08-6', name: 'MUHAMMED UNAIS T', class: 'C2C' },
    { id: 'w-a08-7', name: 'MOHAMMED ISSAM PN', class: 'C2C' },
    { id: 'w-a08-8', name: 'MUHAMMED MUHSIN', class: 'C2C' }
  ],
  'WAVES_A09': [
    { id: 'w-a09-1', name: 'ABDULLA KASSIM KURUNGOT', class: 'C2C' },
    { id: 'w-a09-2', name: 'AZAH OMAR', class: 'C2C' },
    { id: 'w-a09-3', name: 'MUHAMMED SINAN K', class: 'C2B' },
    { id: 'w-a09-4', name: 'AZEEM AHAMMED YM', class: 'C2C' },
    { id: 'w-a09-5', name: 'MUHAMMED JASWIN KP', class: 'C2B' },
    { id: 'w-a09-6', name: 'MUHAMMED NAZAL V', class: 'C2C' },
    { id: 'w-a09-7', name: 'SAVAD MUHAMMED', class: 'C2C' },
    { id: 'w-a09-8', name: 'MUHAMMED SAHEER', class: 'C2B' }
  ],
  'WAVES_A11': [
    { id: 'w-a11-1', name: 'AFHAM SHAHEER NELLIKKA', class: 'C2C' },
    { id: 'w-a11-2', name: 'MUHAMMED JASIL', class: 'C2B' },
    { id: 'w-a11-3', name: 'MUHAMMED JISAL', class: 'C2B' },
    { id: 'w-a11-4', name: 'MUHAMMED K', class: 'C2B' },
    { id: 'w-a11-5', name: 'MUHAMMED ISMAIL', class: 'C2B' },
    { id: 'w-a11-6', name: 'MUHAMMED RASHID', class: 'C2C' },
    { id: 'w-a11-7', name: 'MUHAMMED SUFIYAN K', class: 'C2B' },
    { id: 'w-a11-8', name: 'K MUHAMMED FADHIL', class: 'C2C' }
  ],
  'WAVES_A13': [
    { id: 'w-a13-1', name: 'MOHAMED HADI NAZAL TP', class: 'C2C' },
    { id: 'w-a13-2', name: 'MUHAMMED FIZAN KM', class: 'C2C' },
    { id: 'w-a13-3', name: 'MUHAMMED SINAN M', class: 'C2B' },
    { id: 'w-a13-4', name: 'MUHAMMED RASIL P', class: 'C2C' },
    { id: 'w-a13-5', name: 'MOHAMMED IRFAN', class: 'C2B' },
    { id: 'w-a13-6', name: 'ASHIQ ABDULLA', class: 'C2C' },
    { id: 'w-a13-7', name: 'JAUHAR SHAN', class: 'C2C' }
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

  // Instant splash screen dismiss for zero-delay app opening
  useEffect(() => {
    if (showSplash) {
      setShowSplash(false);
    }
  }, [showSplash]);

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

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 1000);

    fetch('/api/auth/verify', {
      headers: { Authorization: `Bearer ${token}` },
      signal: controller.signal
    })
      .then(r => r.ok ? r.json() : null)
      .then(data => {
        if (data?.user) setCurrentUser(data.user);
        else { localStorage.removeItem('caliph_token'); localStorage.removeItem('caliph_user'); setCurrentUser(null); }
      })
      .catch(() => { })
      .finally(() => {
        clearTimeout(timeoutId);
        setAuthLoading(false);
      });
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
    if (perms.includes('all') || perms.includes(perm)) return true;
    if (perm === 'user_sheet' && perms.includes('admin_sheet')) return true;
    if (perm === 'phone_pass_issue' && perms.includes('admin_phone_pass')) return true;
    if (perm === 'phone_pass' && (perms.includes('admin_phone_pass') || perms.includes('phone_pass_issue'))) return true;
    return false;
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
      try { return JSON.parse(saved); } catch (e) { }
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

  // Class Mentors state & handlers
  const [mentorTypeTab, setMentorTypeTab] = useState('room'); // 'room' | 'class'
  const [classMentors, setClassMentors] = useState(() => {
    const saved = localStorage.getItem('student_tracker_class_mentors');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { }
    }
    return [];
  });

  useEffect(() => {
    localStorage.setItem('student_tracker_class_mentors', JSON.stringify(classMentors));
  }, [classMentors]);

  const [selectedClassMentorIndex, setSelectedClassMentorIndex] = useState(null);
  const [showAddClassMentorModal, setShowAddClassMentorModal] = useState(false);
  const [newClassMentor, setNewClassMentor] = useState({ name: '', classAssigned: '' });
  const [showDeleteClassMentorModal, setShowDeleteClassMentorModal] = useState(false);
  const [selectedClassMentorForDelete, setSelectedClassMentorForDelete] = useState(null);
  const [classMentorSearch, setClassMentorSearch] = useState('');

  const handleAddClassMentor = (e) => {
    e.preventDefault();
    if (!newClassMentor.name.trim() || !newClassMentor.classAssigned.trim()) return;
    const newCM = {
      id: 'cm-' + Date.now(),
      name: newClassMentor.name.trim(),
      classAssigned: newClassMentor.classAssigned.trim(),
      createdAt: new Date().toISOString().split('T')[0]
    };
    setClassMentors([...classMentors, newCM]);
    setNewClassMentor({ name: '', classAssigned: '' });
    setShowAddClassMentorModal(false);
  };

  const handleDeleteClassMentor = (e) => {
    e.preventDefault();
    if (!selectedClassMentorForDelete) return;
    setClassMentors(classMentors.filter(cm => cm.id !== selectedClassMentorForDelete));
    setSelectedClassMentorForDelete(null);
    setShowDeleteClassMentorModal(false);
  };

  const handleDownloadClassMentorSheet = (classMentor) => {
    const classStus = students.filter(s => (s.class || '').toLowerCase() === (classMentor.classAssigned || '').toLowerCase());
    const data = classStus.map((s, idx) => ({
      'Sl No': idx + 1,
      'Student Name': s.name,
      'Class': (s.class || '').toUpperCase(),
      'Room Number': s.roomNumber || 'N/A',
      'Hostel Block': s.hostelBlock || 'N/A',
      'Tally Count': s.tally || 0,
      'Total Fine (₹)': s.fine || 0,
      'Status': s.ineligible ? `Ineligible (${s.ineligibleReason || ''})` : 'Active'
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Class_${classMentor.classAssigned.toUpperCase()}`);
    XLSX.writeFile(wb, `Class_Mentor_${classMentor.name}_${classMentor.classAssigned.toUpperCase()}.xlsx`);
  };

  const handleGenerateClassMentorIR = (classMentor) => {
    const classStus = students.filter(s => (s.class || '').toLowerCase() === (classMentor.classAssigned || '').toLowerCase());
    setIrAssignedStudents(classStus);
    setSelectedIRStudentIds(classStus.map(s => s.id));
    setSelectedIRStudent('ALL');
    setDateModalNextAction('CLASS_MENTOR_IR');
    setShowIRDateModal(true);
  };
  const [performanceView, setPerformanceView] = useState(null); // 'neat', 'room', 'ineligible', null
  const [selectedHostel, setSelectedHostel] = useState(null); // 'MAIN BLOCK BOYS', 'BOYS WAVES', 'COVE', 'HIVE'
  const [selectedRoom, setSelectedRoom] = useState(null); // 'Room 1' .. 'Room 25'
  const DEFAULT_HOSTEL_BLOCKS = [
    { id: 1, name: 'MAIN BLOCK BOYS', icon: '👦', desc: 'Main Boys Hostel' },
    { id: 2, name: 'BOYS WAVES', icon: '🌊', desc: 'Boys Waves Hostel' },
    { id: 3, name: 'COVE', icon: '👧', desc: 'Main Girls Hostel' },
    { id: 4, name: 'HIVE', icon: '🏠', desc: 'Hive Hostel' }
  ];
  const [hostelBlocks, setHostelBlocks] = useState(() => {
    const saved = localStorage.getItem('student_tracker_hostel_blocks');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          return parsed.map(h => {
            if (h.id === 1 || h.name === 'MAIN BLOCK BOYS') return { ...h, name: 'MAIN BLOCK BOYS', icon: '👦', desc: 'Main Boys Hostel' };
            if (h.id === 2 || h.name === 'WAVES' || h.name === 'NEW BLOCK BOYS') return { ...h, name: 'BOYS WAVES', icon: '🌊', desc: 'Boys Waves Hostel' };
            if (h.id === 3 || h.name === 'MAIN BLOCK GIRLS' || h.name === 'GIRLS COVE') return { ...h, name: 'COVE', icon: '👧', desc: 'Main Girls Hostel' };
            if (h.id === 4 || h.name === 'NEW BLOCK GIRLS') return { ...h, name: 'HIVE', icon: '🏠', desc: 'Hive Hostel' };
            return h;
          });
        }
      } catch (e) { console.error(e); }
    }
    return DEFAULT_HOSTEL_BLOCKS;
  });

  useEffect(() => {
    localStorage.setItem('student_tracker_hostel_blocks', JSON.stringify(hostelBlocks));
  }, [hostelBlocks]);

  const [performanceSelectedClass, setPerformanceSelectedClass] = useState(null);
  const [performanceSelectedStudents, setPerformanceSelectedStudents] = useState([]);
  const [showPerformanceSubmitModal, setShowPerformanceSubmitModal] = useState(false);
  const [performanceSubmitData, setPerformanceSubmitData] = useState({ count: 1, reason: '', type: 'tally' });
  const [showPerformanceReasonDropdown, setShowPerformanceReasonDropdown] = useState(false);
  const [showWhatsappModal, setShowWhatsappModal] = useState(false);
  const [whatsappMessage, setWhatsappMessage] = useState('');

  // Room Tally & Fine workflow states
  const [showRoomTallyModal, setShowRoomTallyModal] = useState(false);
  const [roomTallyCount, setRoomTallyCount] = useState(3);
  const [roomTallyReason, setRoomTallyReason] = useState('');
  const [roomTallySummary, setRoomTallySummary] = useState(null);

  const [showRoomFineModal, setShowRoomFineModal] = useState(false);
  const [roomFineAmount, setRoomFineAmount] = useState('');
  const [roomFineReason, setRoomFineReason] = useState('');
  const [roomFineSummary, setRoomFineSummary] = useState(null);

  const [showAddRoomStudentModal, setShowAddRoomStudentModal] = useState(false);
  const [showRemoveRoomStudentModal, setShowRemoveRoomStudentModal] = useState(false);
  const [addRoomStudentName, setAddRoomStudentName] = useState('');
  const [addRoomStudentClass, setAddRoomStudentClass] = useState('S1B');
  const [addRoomStudentModalTab, setAddRoomStudentModalTab] = useState('single'); // 'single' | 'bulk'
  const [bulkImportText, setBulkImportText] = useState('');
  const [bulkParsedStudents, setBulkParsedStudents] = useState([]);
  const [bulkImportFileName, setBulkImportFileName] = useState('');

  // Room Student Allocation State & Persistence
  const [roomStudentMapping, setRoomStudentMapping] = useState(() => {
    const saved = localStorage.getItem('student_tracker_room_mapping');
    let base = { ...ROOM_STUDENT_MAPPING };
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') {
          base = { ...ROOM_STUDENT_MAPPING, ...parsed };
        }
      } catch (e) { console.error(e); }
    }

    // Ensure WAVES room student entries exist
    const wavesKeys = ['WAVES_A02', 'WAVES_A03', 'WAVES_A04', 'WAVES_A05', 'WAVES_A06', 'WAVES_A08', 'WAVES_A09', 'WAVES_A10', 'WAVES_A11', 'WAVES_A13'];
    wavesKeys.forEach(k => {
      if (!base[k] || base[k].length === 0) {
        base[k] = ROOM_STUDENT_MAPPING[k] || [];
      }
    });

    return base;
  });

  useEffect(() => {
    localStorage.setItem('student_tracker_room_mapping', JSON.stringify(roomStudentMapping));
  }, [roomStudentMapping]);

  const WAVES_ROOMS = ['A02', 'A03', 'A04', 'A05', 'A06', 'A08', 'A09', 'A10', 'A11', 'A13'];

  // Dynamic Room List Management per Hostel Block & Persistence
  const DEFAULT_HOSTEL_BLOCK_ROOMS = {
    'MAIN BLOCK BOYS': ['19', '20', '21', '22', '124', '125', '126', '127', '128', '108', '109', '110', '215', '216', '217', '231', '232', '233', '234', '235'],
    'BOYS WAVES': WAVES_ROOMS,
    'COVE': [],
    'HIVE': []
  };

  const [hostelBlockRooms, setHostelBlockRooms] = useState(() => {
    try { localStorage.removeItem('student_tracker_room_list'); } catch (e) { }

    const saved = localStorage.getItem('student_tracker_hostel_block_rooms');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') {
          return {
            'MAIN BLOCK BOYS': parsed['MAIN BLOCK BOYS'] || DEFAULT_HOSTEL_BLOCK_ROOMS['MAIN BLOCK BOYS'],
            'BOYS WAVES': parsed['BOYS WAVES'] || parsed['WAVES'] || parsed['NEW BLOCK BOYS'] || WAVES_ROOMS,
            'COVE': parsed['COVE'] || parsed['MAIN BLOCK GIRLS'] || parsed['GIRLS COVE'] || [],
            'HIVE': parsed['HIVE'] || parsed['NEW BLOCK GIRLS'] || []
          };
        }
      } catch (e) { console.error(e); }
    }
    return DEFAULT_HOSTEL_BLOCK_ROOMS;
  });

  // Flag to ensure local initial state doesn't overwrite DB config before first fetch completes
  const isHostelConfigLoadedRef = useRef(false);

  // --- Synchronize Hostel & Room Config across ALL devices via Backend API ---
  const saveHostelConfigToBackend = (payload) => {
    fetch('/api/hostel-config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).catch(err => console.error("Error saving hostel config to backend:", err));
  };

  useEffect(() => {
    // Initial fetch and periodic sync from backend PostgreSQL DB
    const fetchBackendHostelConfig = () => {
      fetch('/api/hostel-config')
        .then(res => res.ok ? res.json() : null)
        .then(config => {
          if (!config) return;
          if (config.hostel_blocks && Array.isArray(config.hostel_blocks) && config.hostel_blocks.length > 0) {
            setHostelBlocks(config.hostel_blocks);
            localStorage.setItem('student_tracker_hostel_blocks', JSON.stringify(config.hostel_blocks));
          }
          if (config.hostel_block_rooms && typeof config.hostel_block_rooms === 'object') {
            setHostelBlockRooms(prev => ({ ...prev, ...config.hostel_block_rooms }));
            localStorage.setItem('student_tracker_hostel_block_rooms', JSON.stringify(config.hostel_block_rooms));
          }
          if (config.room_student_mapping && typeof config.room_student_mapping === 'object') {
            setRoomStudentMapping(prev => ({ ...prev, ...config.room_student_mapping }));
            localStorage.setItem('student_tracker_room_mapping', JSON.stringify(config.room_student_mapping));
          }
          isHostelConfigLoadedRef.current = true;
        })
        .catch(err => console.error("Error fetching hostel config from backend:", err));
    };

    fetchBackendHostelConfig();
    const interval = setInterval(fetchBackendHostelConfig, 3000); // Sync every 3 seconds for fast multi-user updates
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    localStorage.setItem('student_tracker_hostel_blocks', JSON.stringify(hostelBlocks));
    if (isHostelConfigLoadedRef.current) {
      saveHostelConfigToBackend({ hostel_blocks: hostelBlocks });
    }
  }, [hostelBlocks]);

  useEffect(() => {
    localStorage.setItem('student_tracker_room_mapping', JSON.stringify(roomStudentMapping));
    if (isHostelConfigLoadedRef.current) {
      saveHostelConfigToBackend({ room_student_mapping: roomStudentMapping });
    }
  }, [roomStudentMapping]);

  useEffect(() => {
    localStorage.setItem('student_tracker_hostel_block_rooms', JSON.stringify(hostelBlockRooms));
    if (isHostelConfigLoadedRef.current) {
      saveHostelConfigToBackend({ hostel_block_rooms: hostelBlockRooms });
    }
  }, [hostelBlockRooms]);

  // Active Hostel Block Name and Room List
  const activeHostelName = selectedHostel || 'MAIN BLOCK BOYS';
  const hostelRoomList = hostelBlockRooms[activeHostelName] || [];

  // Helper for room keys in roomStudentMapping:
  const getHostelRoomKey = (hostelName, roomNum) => {
    const h = hostelName || activeHostelName || 'MAIN BLOCK BOYS';
    const r = String(roomNum).replace(/room\s*/i, '').trim();
    return `${h}_${r}`;
  };

  const getHostelRoomStudents = (hostelName, roomNum) => {
    const h = hostelName || activeHostelName || 'MAIN BLOCK BOYS';
    const rKey = String(roomNum).replace(/room\s*/i, '').trim();
    const fullKey = `${h}_${rKey}`;

    let rawList = [];
    if (roomStudentMapping[fullKey] !== undefined) {
      rawList = roomStudentMapping[fullKey];
    } else if (roomStudentMapping[`BOYS WAVES_${rKey}`] !== undefined) {
      rawList = roomStudentMapping[`BOYS WAVES_${rKey}`];
    } else if (roomStudentMapping[`WAVES_${rKey}`] !== undefined) {
      rawList = roomStudentMapping[`WAVES_${rKey}`];
    } else if (roomStudentMapping[`COVE_${rKey}`] !== undefined) {
      rawList = roomStudentMapping[`COVE_${rKey}`];
    } else if (roomStudentMapping[`MAIN BLOCK BOYS_${rKey}`] !== undefined) {
      rawList = roomStudentMapping[`MAIN BLOCK BOYS_${rKey}`];
    } else if (roomStudentMapping[`HIVE_${rKey}`] !== undefined) {
      rawList = roomStudentMapping[`HIVE_${rKey}`];
    } else if (roomStudentMapping[`NEW BLOCK GIRLS_${rKey}`] !== undefined) {
      rawList = roomStudentMapping[`NEW BLOCK GIRLS_${rKey}`];
    } else if ((h === 'MAIN BLOCK BOYS' || h === 'COVE') && roomStudentMapping[rKey] !== undefined) {
      rawList = roomStudentMapping[rKey];
    } else {
      rawList = students.filter(s => String(s.room || s.roomNumber || '').trim() === rKey);
    }

    // Auto-link every room student in all rooms & hostel blocks to their database record
    return rawList.map(item => {
      const dbMatch = students.find(s =>
        String(s.id) === String(item.id) ||
        (s.name && item.name && s.name.trim().toLowerCase() === item.name.trim().toLowerCase())
      );
      if (dbMatch) {
        return {
          ...dbMatch,
          room: rKey,
          roomNumber: rKey,
          hostelBlock: h
        };
      }
      return item;
    });
  };


  // Handler to add student to current open room
  const handleAddStudentToPerformanceRoom = (e) => {
    e.preventDefault();
    if (!addRoomStudentName.trim()) return;
    const currentRoomNum = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
    if (!currentRoomNum) return;

    const fullRoomKey = getHostelRoomKey(activeHostelName, currentRoomNum);

    const newStudentObj = {
      id: 'r' + currentRoomNum + '-' + Date.now(),
      name: addRoomStudentName.trim().toUpperCase(),
      class: addRoomStudentClass.trim().toUpperCase()
    };

    setRoomStudentMapping(prev => {
      const existing = prev[fullRoomKey] || prev[currentRoomNum] || [];
      if (existing.some(s => s.name.toUpperCase() === newStudentObj.name && s.class.toUpperCase() === newStudentObj.class)) {
        alert(`${newStudentObj.name} is already in ${selectedRoom}!`);
        return prev;
      }
      return {
        ...prev,
        [fullRoomKey]: [...existing, newStudentObj]
      };
    });

    setAddRoomStudentName('');
    setShowAddRoomStudentModal(false);
  };

  // Helper to parse student lines e.g. "Nitha Fathima (C2A)" or tab/comma separated
  const parseBulkStudentLine = (line, fallbackClass = 'S1B') => {
    if (!line || !line.trim()) return null;
    const str = line.trim();

    // Pattern 1: "Nitha Fathima (C2A)" or "Nitha Fathima (S2A)" or "Nitha Fathima [C2A]"
    const parenMatch = str.match(/^(.+?)\s*[\(\[]\s*(?:class:?\s*)?([A-Za-z0-9]+)\s*[\)\]]$/i);
    if (parenMatch) {
      return {
        name: parenMatch[1].trim().toUpperCase(),
        class: parenMatch[2].trim().toUpperCase()
      };
    }

    // Pattern 2: Tab, Comma, Semicolon, Pipe separated: "Nitha Fathima, C2A" or "Nitha Fathima\tC2A"
    if (/[\t,;|]/.test(str)) {
      const parts = str.split(/[\t,;|]/).map(s => s.trim()).filter(Boolean);
      if (parts.length >= 2) {
        if (/^(name|student|student name|full name)$/i.test(parts[0]) && /^(class|grade|sec|section)$/i.test(parts[1])) {
          return null; // header row
        }
        return {
          name: parts[0].toUpperCase(),
          class: parts[1].replace(/^class\s*/i, '').trim().toUpperCase()
        };
      }
    }

    // Pattern 3: Space separated where last token looks like a class code e.g. "Nitha Fathima C2A"
    const spaceParts = str.split(/\s+/);
    if (spaceParts.length >= 2) {
      const lastPart = spaceParts[spaceParts.length - 1];
      if (/^[A-Za-z]?[0-9]+[A-Za-z]?$/i.test(lastPart) || /^[A-Za-z]+[0-9]+[A-Za-z]?$/i.test(lastPart)) {
        const name = spaceParts.slice(0, spaceParts.length - 1).join(' ').trim().toUpperCase();
        const cls = lastPart.trim().toUpperCase();
        return { name, class: cls };
      }
    }

    // Fallback: whole line as student name, fallback class
    return {
      name: str.toUpperCase(),
      class: fallbackClass || 'S1B'
    };
  };

  const handleBulkTextChange = (text) => {
    setBulkImportText(text);
    if (!text.trim()) {
      setBulkParsedStudents([]);
      return;
    }
    const lines = text.split(/\r?\n/);
    const parsed = [];
    lines.forEach(l => {
      const res = parseBulkStudentLine(l, addRoomStudentClass);
      if (res && res.name) parsed.push(res);
    });
    setBulkParsedStudents(parsed);
  };

  const handleBulkExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkImportFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const parsed = [];

        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
          if (!rows || rows.length === 0) return;

          let nameCol = -1;
          let classCol = -1;
          let startRow = 0;

          const firstRow = rows[0];
          if (Array.isArray(firstRow)) {
            firstRow.forEach((h, i) => {
              if (!h) return;
              const s = String(h).toLowerCase().trim();
              if (s === 'name' || s === 'student name' || s === 'student') nameCol = i;
              if (s === 'class' || s === 'grade' || s === 'sec') classCol = i;
            });

            if (nameCol !== -1 || classCol !== -1) {
              startRow = 1;
            } else {
              nameCol = 0;
              classCol = 1;
            }

            for (let i = startRow; i < rows.length; i++) {
              const r = rows[i];
              if (!r || r.length === 0) continue;

              const val0 = r[nameCol !== -1 ? nameCol : 0];
              const val1 = r[classCol !== -1 ? classCol : 1];

              if (val0 && !val1) {
                const res = parseBulkStudentLine(String(val0), addRoomStudentClass);
                if (res && res.name) parsed.push(res);
              } else if (val0) {
                parsed.push({
                  name: String(val0).trim().toUpperCase(),
                  class: String(val1 || addRoomStudentClass || 'S1B').replace(/^class\s*/i, '').trim().toUpperCase()
                });
              }
            }
          }
        });

        setBulkParsedStudents(parsed);
        const formattedText = parsed.map(s => `${s.name} (${s.class})`).join('\n');
        setBulkImportText(formattedText);
      } catch (err) {
        alert("Error reading Excel file: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleBulkImportStudentsToRoom = (e) => {
    if (e) e.preventDefault();
    if (bulkParsedStudents.length === 0) return;

    const currentRoomNum = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
    if (!currentRoomNum) return;

    const fullRoomKey = getHostelRoomKey(activeHostelName, currentRoomNum);
    const now = Date.now();

    setRoomStudentMapping(prev => {
      const existing = prev[fullRoomKey] || prev[currentRoomNum] || [];
      const existingNames = new Set(existing.map(s => s.name.toUpperCase()));

      const newAdditions = [];
      const newClassesToRegister = new Set();

      bulkParsedStudents.forEach((s, idx) => {
        const cleanName = s.name.trim().toUpperCase();
        const cleanClass = s.class.trim().toUpperCase() || 'S1B';
        if (!cleanName) return;

        if (!existingNames.has(cleanName)) {
          newAdditions.push({
            id: 'r' + currentRoomNum + '-' + (now + idx),
            name: cleanName,
            class: cleanClass
          });
          existingNames.add(cleanName);
          if (cleanClass) newClassesToRegister.add(cleanClass.toLowerCase());
        }
      });

      if (newClassesToRegister.size > 0) {
        setCLASSES(prevCls => {
          const updated = [...prevCls];
          newClassesToRegister.forEach(c => {
            if (!updated.includes(c)) updated.push(c);
          });
          localStorage.setItem('caliph_classes', JSON.stringify(updated));
          return updated;
        });
      }

      if (newAdditions.length === 0) {
        alert(`All parsed students are already in ${selectedRoom}!`);
        return prev;
      }

      alert(`Successfully added ${newAdditions.length} student(s) to ${selectedRoom}!`);

      return {
        ...prev,
        [fullRoomKey]: [...existing, ...newAdditions]
      };
    });

    setBulkImportText('');
    setBulkParsedStudents([]);
    setBulkImportFileName('');
    setShowAddRoomStudentModal(false);
  };

  // Handler to remove students from current open room
  const handleRemoveStudentsFromPerformanceRoom = (idsToRemove) => {
    if (!idsToRemove || idsToRemove.length === 0) return;
    const currentRoomNum = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
    if (!currentRoomNum) return;

    const fullRoomKey = getHostelRoomKey(activeHostelName, currentRoomNum);

    setRoomStudentMapping(prev => {
      const existing = prev[fullRoomKey] || prev[currentRoomNum] || [];
      const updatedList = existing.filter(s => !idsToRemove.includes(s.id));
      return {
        ...prev,
        [fullRoomKey]: updatedList
      };
    });

    setPerformanceSelectedStudents(prev => prev.filter(id => !idsToRemove.includes(id)));
    setRemoveSelectedRoomStudentIds([]);
    setShowRemoveRoomStudentModal(false);
  };

  // Modal & Edit States for Rooms
  const [showAddNewRoomModal, setShowAddNewRoomModal] = useState(false);
  const [newRoomNumberInput, setNewRoomNumberInput] = useState('');

  const [showDeleteRoomModal, setShowDeleteRoomModal] = useState(false);
  const [roomsToDelete, setRoomsToDelete] = useState([]);

  const [showEditRoomNumberModal, setShowEditRoomNumberModal] = useState(false);
  const [editingRoomOldNum, setEditingRoomOldNum] = useState('');
  const [editingRoomNewNum, setEditingRoomNewNum] = useState('');

  // Hostel Block Rename state & handlers
  const [showEditHostelModal, setShowEditHostelModal] = useState(false);
  const [editingHostelId, setEditingHostelId] = useState(null);
  const [editingHostelOldName, setEditingHostelOldName] = useState('');
  const [editingHostelNewName, setEditingHostelNewName] = useState('');

  const openEditHostelModal = (hostel) => {
    setEditingHostelId(hostel.id);
    setEditingHostelOldName(hostel.name);
    setEditingHostelNewName(hostel.name);
    setShowEditHostelModal(true);
  };

  const handleEditHostelNameSubmit = (e) => {
    if (e) e.preventDefault();
    const oldName = editingHostelOldName.trim();
    const newName = editingHostelNewName.trim().toUpperCase();

    if (!newName) {
      alert("Hostel name cannot be empty.");
      return;
    }

    if (oldName === newName) {
      setShowEditHostelModal(false);
      return;
    }

    if (hostelBlocks.some(h => h.id !== editingHostelId && h.name.toUpperCase() === newName)) {
      alert(`Hostel "${newName}" already exists!`);
      return;
    }

    // Update hostelBlocks list
    setHostelBlocks(prev => prev.map(h => h.id === editingHostelId ? { ...h, name: newName } : h));

    // Update hostelBlockRooms keys
    setHostelBlockRooms(prev => {
      const next = { ...prev };
      if (next[oldName]) {
        next[newName] = next[oldName];
        delete next[oldName];
      } else if (!next[newName]) {
        next[newName] = [];
      }
      return next;
    });

    // Update roomStudentMapping keys
    setRoomStudentMapping(prev => {
      const next = { ...prev };
      Object.keys(prev).forEach(key => {
        if (key.startsWith(`${oldName}_`)) {
          const newKey = key.replace(`${oldName}_`, `${newName}_`);
          next[newKey] = prev[key];
          delete next[key];
        }
      });
      return next;
    });

    // If active selected Hostel is oldName, update selectedHostel
    if (selectedHostel === oldName) {
      setSelectedHostel(newName);
    }

    setShowEditHostelModal(false);
  };

  // Long press timer ref
  const roomLongPressTimerRef = useRef(null);
  const isLongPressTriggeredRef = useRef(false);

  const handleRoomPressStart = (num) => {
    isLongPressTriggeredRef.current = false;
    roomLongPressTimerRef.current = setTimeout(() => {
      isLongPressTriggeredRef.current = true;
      setEditingRoomOldNum(String(num));
      setEditingRoomNewNum(String(num));
      setShowEditRoomNumberModal(true);
    }, 500);
  };

  const handleRoomPressEnd = () => {
    if (roomLongPressTimerRef.current) {
      clearTimeout(roomLongPressTimerRef.current);
      roomLongPressTimerRef.current = null;
    }
  };

  const handleAddNewRoom = (e) => {
    e.preventDefault();
    const roomNumStr = newRoomNumberInput.trim();
    if (!roomNumStr) return;

    const currentRooms = hostelBlockRooms[activeHostelName] || [];
    if (currentRooms.includes(roomNumStr)) {
      alert(`Room ${roomNumStr} already exists in ${activeHostelName}!`);
      return;
    }

    const updatedRooms = {
      ...hostelBlockRooms,
      [activeHostelName]: [...(hostelBlockRooms[activeHostelName] || []), roomNumStr]
    };

    setHostelBlockRooms(updatedRooms);
    saveHostelConfigToBackend({ hostel_block_rooms: updatedRooms });

    setNewRoomNumberInput('');
    setShowAddNewRoomModal(false);
  };

  const handleDeleteRooms = () => {
    if (roomsToDelete.length === 0) return;
    const updatedRooms = {
      ...hostelBlockRooms,
      [activeHostelName]: (hostelBlockRooms[activeHostelName] || []).filter(num => !roomsToDelete.includes(num))
    };

    setHostelBlockRooms(updatedRooms);
    saveHostelConfigToBackend({ hostel_block_rooms: updatedRooms });

    setRoomsToDelete([]);
    setShowDeleteRoomModal(false);
  };

  const handleEditRoomNumberSubmit = (e) => {
    e.preventDefault();
    const oldNum = editingRoomOldNum.trim();
    const newNum = editingRoomNewNum.trim();
    if (!newNum || oldNum === newNum) {
      setShowEditRoomNumberModal(false);
      return;
    }

    const currentRooms = hostelBlockRooms[activeHostelName] || [];
    if (currentRooms.includes(newNum) && newNum !== oldNum) {
      alert(`Room ${newNum} already exists in ${activeHostelName}!`);
      return;
    }

    // Update room list for active hostel block
    const updatedRooms = {
      ...hostelBlockRooms,
      [activeHostelName]: (hostelBlockRooms[activeHostelName] || []).map(num => num === oldNum ? newNum : num)
    };
    setHostelBlockRooms(updatedRooms);
    saveHostelConfigToBackend({ hostel_block_rooms: updatedRooms });

    // Update room student mapping
    const oldKey = getHostelRoomKey(activeHostelName, oldNum);
    const newKey = getHostelRoomKey(activeHostelName, newNum);

    setRoomStudentMapping(prev => {
      const updated = { ...prev };
      if (updated[oldKey] !== undefined) {
        updated[newKey] = updated[oldKey];
        delete updated[oldKey];
      }
      saveHostelConfigToBackend({ room_student_mapping: updated });
      return updated;
    });

    setShowEditRoomNumberModal(false);
  };

  const [showAddIneligibleModal, setShowAddIneligibleModal] = useState(false);
  const [ineligibleSelectedStudents, setIneligibleSelectedStudents] = useState([]);
  const [ineligibleReasonInput, setIneligibleReasonInput] = useState('');

  const [spotClass, setSpotClass] = useState('');
  const [spotNameSearch, setSpotNameSearch] = useState('');
  const [spotSelectedStudentIds, setSpotSelectedStudentIds] = useState([]);
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
  const [sheetSelectedStudentIds, setSheetSelectedStudentIds] = useState([]);

  // Morning Bliss State
  const [morningBlissClass, setMorningBlissClass] = useState('');
  const [morningBlissNameSearch, setMorningBlissNameSearch] = useState('');
  const [morningBlissTopic, setMorningBlissTopic] = useState('');
  const [morningBlissMark, setMorningBlissMark] = useState('');
  const [morningBlissEv, setMorningBlissEv] = useState('');
  const [showMorningBlissDropdown, setShowMorningBlissDropdown] = useState(false);
  const [morningBlissDuration, setMorningBlissDuration] = useState(0); // milliseconds
  const [isStopwatchRunning, setIsStopwatchRunning] = useState(false);
  const stopwatchStartTimeRef = useRef(null);
  const stopwatchAccumulatedRef = useRef(0);
  const [mbFromDate, setMbFromDate] = useState(() => new Date().toISOString().split('T')[0]);
  const [mbToDate, setMbToDate] = useState(() => new Date().toISOString().split('T')[0]);
  const [selectedMbStudentIds, setSelectedMbStudentIds] = useState([]);

  const handleDeleteSelectedMorningBliss = (targetIds = null) => {
    const ids = targetIds || selectedMbStudentIds;
    if (!ids || ids.length === 0) {
      alert("Please select at least one student record to delete.");
      return;
    }
    const count = ids.length;
    if (window.confirm(`Are you sure you want to delete Morning Bliss record for ${count === 1 ? 'this student' : count + ' selected student(s)'}?`)) {
      const updatedStudents = students.map(s => {
        if (ids.includes(s.id)) {
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
      setSelectedMbStudentIds(prev => prev.filter(id => !ids.includes(id)));

      fetch('/api/students/bulk-upsert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          students: updatedStudents.filter(s => ids.includes(s.id))
        })
      }).catch(err => console.error("Error bulk clearing selected morning bliss:", err));
    }
  };
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
        if (stopwatchStartTimeRef.current) {
          setMorningBlissDuration(Date.now() - stopwatchStartTimeRef.current);
        }
      }, 30);
    }

    const handleSync = () => {
      if (isStopwatchRunning && stopwatchStartTimeRef.current) {
        setMorningBlissDuration(Date.now() - stopwatchStartTimeRef.current);
      }
    };

    document.addEventListener('visibilitychange', handleSync);
    window.addEventListener('focus', handleSync);

    return () => {
      clearInterval(interval);
      document.removeEventListener('visibilitychange', handleSync);
      window.removeEventListener('focus', handleSync);
    };
  }, [isStopwatchRunning]);

  const formatMBTime = (totalMs) => {
    if (!totalMs || isNaN(totalMs)) return '00:00:00';
    const totalSeconds = Math.floor(totalMs / 1000);
    const m = Math.floor(totalSeconds / 60).toString().padStart(2, '0');
    const s = (totalSeconds % 60).toString().padStart(2, '0');
    const ms = Math.floor((totalMs % 1000) / 10).toString().padStart(2, '0');
    return `${m}:${s}:${ms}`;
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
    setIsStopwatchRunning(false);
    stopwatchAccumulatedRef.current = 0;
    stopwatchStartTimeRef.current = null;
    setMorningBlissDuration(0);
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

    let targetStudentIds = [...sheetSelectedStudentIds];

    if (targetStudentIds.length === 0 && sheetNameSearch.trim()) {
      const classMatch = (s) => (!sheetClass || sheetClass === 'ALL') ? true : s.class === sheetClass;
      const match = students.find(s => classMatch(s) && s.name.toLowerCase() === sheetNameSearch.trim().toLowerCase());
      const partial = match || students.find(s => classMatch(s) && s.name.toLowerCase().startsWith(sheetNameSearch.trim().toLowerCase()));
      if (partial) {
        targetStudentIds = [partial.id];
      }
    }

    if (!sheetClass || targetStudentIds.length === 0 || !sheetReason || !performanceView) {
      if (targetStudentIds.length === 0) {
        alert("Please select at least one student!");
      }
      return;
    }

    let typeLabel = '';
    if (performanceView === 'sheets_black') typeLabel = 'Black Sheet';
    if (performanceView === 'sheets_yellow') typeLabel = 'Yellow Sheet';
    if (performanceView === 'sheets_apology') typeLabel = 'Apology Sheet';

    let studentsToUpsert = [];
    let updatedStudentsList = [...students];
    let selectedStudentSummaries = [];

    targetStudentIds.forEach(studentId => {
      const studentToUpdate = updatedStudentsList.find(s => s.id === studentId);
      if (studentToUpdate) {
        let updated = { ...studentToUpdate };
        if (performanceView === 'sheets_black') {
          updated.sheetTally = (studentToUpdate.sheetTally || 0) - 8;
          updated.ineligible = true;
          updated.ineligibleReason = sheetReason ? `Black Sheet - ${sheetReason}` : 'Black Sheet';
          logHistory(studentToUpdate.id, 'Black Sheet', -8, sheetReason);
        } else if (performanceView === 'sheets_yellow') {
          updated.sheetTally = (studentToUpdate.sheetTally || 0) - 4;
          logHistory(studentToUpdate.id, 'Yellow Sheet', -4, sheetReason);
        } else if (performanceView === 'sheets_apology') {
          updated.sheetTally = (studentToUpdate.sheetTally || 0) - 1.5;
          logHistory(studentToUpdate.id, 'Apology Sheet', -1.5, sheetReason);
        }
        studentsToUpsert.push(updated);
        selectedStudentSummaries.push(`- ${studentToUpdate.name} (${studentToUpdate.class.toUpperCase()})`);
        updatedStudentsList = updatedStudentsList.map(s => s.id === studentId ? updated : s);
      }
    });

    setStudents(updatedStudentsList);

    fetch('/api/students/bulk-upsert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ students: studentsToUpsert })
    }).catch(err => console.error("Error bulk upserting sheet records:", err));

    setSheetNameSearch('');
    setSheetSelectedStudentIds([]);
    setSheetReason('');
    setShowSheetDropdown(false);

    const now = new Date();
    const formattedDate = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    const formattedTime = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const classDisplay = sheetClass === 'ALL' ? 'ALL CLASSES' : sheetClass.toUpperCase();

    const sheetSummary = `📄 *SHEET REPORT*
Type: *${typeLabel}*
Date: ${formattedDate}
Time: ${formattedTime}
Class: *${classDisplay}*
Total Students: *${selectedStudentSummaries.length}*
Reason: ${sheetReason}

Students:
${selectedStudentSummaries.join('\n')}`;

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
    const programStarSummary = `*${programClass.toUpperCase()}*

Name: ${studentToUpdate.name}
Programme: ${programReason || 'N/A'}
Star Count: ${amount}star ⭐`;

    setSummaryText(programStarSummary);
    setShowSummaryModal(true);
  };

  const handleSpotFineSubmit = (e) => {
    e.preventDefault();

    // Determine target students
    let targetStudentIds = [...spotSelectedStudentIds];

    // Fallback: If user typed a search name but didn't explicitly check a box, find matching student
    if (targetStudentIds.length === 0 && spotNameSearch.trim()) {
      const classMatch = (s) => !spotClass || spotClass === 'ALL' ? true : s.class === spotClass;
      const match = students.find(s => classMatch(s) && s.name.toLowerCase() === spotNameSearch.trim().toLowerCase());
      const partial = match || students.find(s => classMatch(s) && s.name.toLowerCase().startsWith(spotNameSearch.trim().toLowerCase()));
      if (partial) {
        targetStudentIds = [partial.id];
      }
    }

    if (!spotClass || targetStudentIds.length === 0 || !spotAmount) {
      if (targetStudentIds.length === 0) {
        alert("Please select at least one student!");
      }
      return;
    }

    const amount = Number(spotAmount);
    let studentsToUpsert = [];
    let updatedStudentsList = [...students];
    let selectedStudentSummaries = [];

    targetStudentIds.forEach(studentId => {
      const studentToUpdate = updatedStudentsList.find(s => s.id === studentId);
      if (studentToUpdate) {
        const updated = {
          ...studentToUpdate,
          fine: (studentToUpdate.fine || 0) + amount,
          fineCount: (studentToUpdate.fineCount || 0) + 1,
          fineReason: spotReason || studentToUpdate.fineReason
        };
        logHistory(studentToUpdate.id, 'Spot Fine', amount, spotReason);
        studentsToUpsert.push(updated);
        selectedStudentSummaries.push(`- ${studentToUpdate.name} (${studentToUpdate.class.toUpperCase()})`);
        updatedStudentsList = updatedStudentsList.map(s => s.id === studentId ? updated : s);
      }
    });

    setStudents(updatedStudentsList);

    fetch('/api/students/bulk-upsert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ students: studentsToUpsert })
    }).catch(err => console.error("Error bulk upserting spot fine:", err));

    setSpotNameSearch('');
    setSpotSelectedStudentIds([]);
    setSpotReason('');
    setSpotAmount('');
    setShowSpotDropdown(false);

    const now = new Date();
    const formattedDate = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    const formattedTime = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const classDisplay = spotClass === 'ALL' ? '*ALL CLASSES*' : `*${spotClass.toUpperCase()}*`;

    const spotFineSummary = `🚨 *SPOT FINE REPORT*
Date: ${formattedDate}
Time: ${formattedTime}
Class: ${classDisplay}
Reason: ${spotReason || 'N/A'}
Fine Amount: ₹${amount}
Total Fined: ${selectedStudentSummaries.length} Student(s)

Fined Students:
${selectedStudentSummaries.join('\n')}`;

    setSummaryText(spotFineSummary);
    setShowSummaryModal(true);
  };


  const handlePerformanceSubmit = (e) => {
    e.preventDefault();
    const amount = Number(performanceSubmitData.count);
    const reason = performanceSubmitData.reason;

    const firstSelectedStudent = students.find(s => performanceSelectedStudents.includes(s.id));
    const targetClassHeader = (performanceSelectedClass || firstSelectedStudent?.class || 'C2A').toUpperCase();
    let messageLines = [targetClassHeader, ''];
    let studentsToUpsert = [];

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
        studentsToUpsert.push(updated);
        logHistory(s.id, isNeat ? 'N&O' : 'tally', isNeat ? 1 : amount, reason);
        const starsStr = isNeat ? '⭐' : '⭐'.repeat(Math.max(1, Math.min(amount, 20)));
        messageLines.push(`${s.name} ${amount}${isNeat ? 'tally' : (performanceSubmitData.type || 'star')} ${starsStr}`);
        return updated;
      }
      return s;
    });

    setStudents(updatedStudents);

    if (studentsToUpsert.length > 0) {
      fetch('/api/students/bulk-upsert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ students: studentsToUpsert })
      }).catch(err => console.error("Error bulk upserting:", err));
    }

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
  const isAdminAuthenticated = isSuperAdmin || currentUser?.role === 'admin' || hasPermission('admin_sheet') || hasPermission('admin_phone_pass');
  const [passwordError, setPasswordError] = useState('');
  // User management state
  const [adminSubTab, setAdminSubTab] = useState('sheet'); // 'sheet' | 'phone_pass' | 'users'
  const [usersList, setUsersList] = useState([]);
  const [usersLoading, setUsersLoading] = useState(false);
  const [showUserModal, setShowUserModal] = useState(false);
  const [editingUser, setEditingUser] = useState(null);
  const [userForm, setUserForm] = useState({ username: '', password: '', role: 'user', permissions: [] });
  const [userFormError, setUserFormError] = useState('');

  useEffect(() => {
    if (currentUser) {
      if (hasPermission('scoring')) {
        setActiveTab('scoring');
      } else if (hasPermission('phone_pass') || hasPermission('phone_pass_issue') || hasPermission('admin_phone_pass')) {
        setActiveTab('phone_pass');
      } else if (hasPermission('mentor')) {
        setActiveTab('mentor');
      } else if (hasPermission('performance')) {
        setActiveTab('performance');
      } else if (hasPermission('user_sheet') || hasPermission('admin_sheet') || hasPermission('admin_phone_pass') || currentUser.role === 'admin' || currentUser.role === 'super_admin') {
        setActiveTab('admin');
      }
    }
  }, [currentUser]);

  useEffect(() => {
    if (activeTab === 'admin') {
      if (!hasPermission('admin_sheet') && hasPermission('admin_phone_pass')) {
        setAdminSubTab('phone_pass');
      }
    }
  }, [activeTab, currentUser]);

  // --- Phone Pass System State & Persistence ---
  const [phonePasses, setPhonePasses] = useState(() => {
    const saved = localStorage.getItem('student_tracker_phone_passes');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { }
    }
    return [];
  });
  const [showIssuePhonePassModal, setShowIssuePhonePassModal] = useState(false);
  const [phonePassStep, setPhonePassStep] = useState(1); // 1: Select Student, 2: Pass Form
  const [phonePassSelectedStudent, setPhonePassSelectedStudent] = useState(null);
  const [phonePassReason, setPhonePassReason] = useState('Calling Home');
  const [phonePassAllowedMins, setPhonePassAllowedMins] = useState(60);
  const [phonePassCustomAllowedTime, setPhonePassCustomAllowedTime] = useState('');
  const [phonePassReturnDate, setPhonePassReturnDate] = useState(() => new Date().toISOString().split('T')[0]);
  const [phonePassSearchQuery, setPhonePassSearchQuery] = useState('');
  const [phonePassFilterStatus, setPhonePassFilterStatus] = useState('all'); // 'all', 'OUT', 'IN', 'late'
  const [phonePassStudentTypeFilter, setPhonePassStudentTypeFilter] = useState('school'); // 'school' | 'home'
  const [phonePassStudentSearch, setPhonePassStudentSearch] = useState('');
  const [phonePassClassFilter, setPhonePassClassFilter] = useState('all');

  // --- Monthly Leave Bulk Pass State ---
  const [showMonthlyLeaveModal, setShowMonthlyLeaveModal] = useState(false);
  const [monthlyLeaveDate, setMonthlyLeaveDate] = useState(() => new Date().toISOString().split('T')[0]);
  const [monthlyLeaveTime, setMonthlyLeaveTime] = useState('18:00');
  const [monthlyLeaveClassFilter, setMonthlyLeaveClassFilter] = useState('all');
  const [monthlyLeaveReason, setMonthlyLeaveReason] = useState('Monthly Leave');

  const handleIssueMonthlyLeavePasses = async (e) => {
    if (e) e.preventDefault();
    if (!monthlyLeaveDate || !monthlyLeaveTime) {
      alert("Please select both a return date and time.");
      return;
    }

    // Combine date and time to create allowedUntil ISO timestamp
    const allowedUntilTs = new Date(`${monthlyLeaveDate}T${monthlyLeaveTime}:00`);
    if (isNaN(allowedUntilTs.getTime())) {
      alert("Invalid date or time selected.");
      return;
    }

    // Filter eligible students (NOT ineligible)
    const eligibleStudents = students.filter(s => {
      if (s.ineligible) return false;
      if (monthlyLeaveClassFilter !== 'all' && s.class.toLowerCase() !== monthlyLeaveClassFilter.toLowerCase()) return false;
      return true;
    });

    if (eligibleStudents.length === 0) {
      alert("No eligible students found matching the selected class filter.");
      return;
    }

    // Exclude students who ALREADY have an active pass ('ISSUED' or 'OUT')
    const studentsToIssue = eligibleStudents.filter(st => {
      const activePass = phonePasses.some(p => String(p.studentId) === String(st.id) && ['ISSUED', 'OUT'].includes(p.status));
      return !activePass;
    });

    if (studentsToIssue.length === 0) {
      alert("All eligible students already have active phone passes!");
      return;
    }

    const now = new Date();
    const batchPasses = studentsToIssue.map(st => ({
      studentId: st.id,
      studentName: st.name,
      studentClass: st.class,
      phoneModel: st.phoneModel || '',
      studentType: st.phoneType || 'school',
      startTime: now.toISOString(),
      allowedUntil: allowedUntilTs.toISOString(),
      reason: monthlyLeaveReason || 'Monthly Leave',
      issuedBy: currentUser?.username || 'Admin'
    }));

    try {
      const res = await fetch('/api/phone-passes/batch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ passes: batchPasses })
      });

      if (res.ok) {
        const data = await res.json();
        alert(`Successfully issued Monthly Leave Phone Passes to ${data.count} eligible student(s)! (Ineligible students skipped)`);
        fetchPhonePasses();
      } else {
        const errData = await res.json().catch(() => ({}));
        alert(errData.error || "Failed to issue batch Monthly Leave passes.");
      }
    } catch (err) {
      console.error("Error issuing Monthly Leave passes:", err);
      // Fallback local update
      setPhonePasses(prev => [...batchPasses.map(p => ({ ...p, id: 'pass-ml-' + p.studentId + '-' + Date.now(), status: 'ISSUED', isLate: false, createdAt: now.toISOString() })), ...prev]);
      alert(`Issued Monthly Leave passes to ${batchPasses.length} eligible student(s).`);
    }

    setShowMonthlyLeaveModal(false);
  };

  const fetchPhonePasses = () => {
    fetch('/api/phone-passes')
      .then(res => res.ok ? res.json() : null)
      .then(data => {
        if (Array.isArray(data)) {
          setPhonePasses(data);
          localStorage.setItem('student_tracker_phone_passes', JSON.stringify(data));
        }
      })
      .catch(err => console.error("Error fetching phone passes:", err));
  };

  useEffect(() => {
    fetchPhonePasses();
    const interval = setInterval(fetchPhonePasses, 8000);
    return () => clearInterval(interval);
  }, []);

  const handleIssuePhonePassSubmit = async (e) => {
    if (e) e.preventDefault();
    if (!phonePassSelectedStudent) return;

    if (!phonePassSelectedStudent.phoneModel || !phonePassSelectedStudent.phoneModel.trim()) {
      alert(`⚠️ Cannot Issue Phone Pass! Student "${phonePassSelectedStudent.name}" has no Phone Name registered. Please add the phone name in Phone Pass Management first.`);
      return;
    }

    // Check if student ALREADY has an active pass (status 'ISSUED' or 'OUT')
    const activePassExists = phonePasses.some(
      p => String(p.studentId) === String(phonePassSelectedStudent.id) && ['ISSUED', 'OUT'].includes(p.status)
    );

    if (activePassExists) {
      alert(`⚠️ Student ${phonePassSelectedStudent.name} already has an active phone pass! Please return the previous pass before issuing a new one.`);
      return;
    }

    const now = new Date();
    let allowedTs;
    const returnDateStr = phonePassReturnDate || now.toISOString().split('T')[0];

    if (phonePassCustomAllowedTime) {
      const [hours, mins] = phonePassCustomAllowedTime.split(':');
      allowedTs = new Date(`${returnDateStr}T${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}:00`);
    } else {
      const calcTime = new Date(now.getTime() + (Number(phonePassAllowedMins) || 60) * 60000);
      const timeString = calcTime.toTimeString().split(' ')[0];
      allowedTs = new Date(`${returnDateStr}T${timeString}`);
    }

    if (isNaN(allowedTs.getTime())) {
      allowedTs = new Date(now.getTime() + (Number(phonePassAllowedMins) || 60) * 60000);
    }

    const passData = {
      id: 'pass-' + Date.now(),
      studentId: phonePassSelectedStudent.id,
      studentName: phonePassSelectedStudent.name,
      studentClass: phonePassSelectedStudent.class,
      phoneModel: phonePassSelectedStudent.phoneModel || 'Standard Phone',
      studentType: phonePassSelectedStudent.phoneType || 'school',
      startTime: now.toISOString(),
      allowedUntil: allowedTs.toISOString(),
      reason: phonePassReason || 'Calling Home',
      issuedBy: currentUser?.username || 'Admin'
    };

    try {
      const res = await fetch('/api/phone-passes', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(passData)
      });

      if (!res.ok) {
        const errData = await res.json();
        alert(errData.error || "Failed to issue phone pass");
        return;
      }

      const createdPass = await res.json();
      setPhonePasses(prev => [createdPass, ...prev]);
      setShowIssuePhonePassModal(false);
      setPhonePassSelectedStudent(null);
      setPhonePassReason('Calling Home');
      setPhonePassCustomAllowedTime('');
      setPhonePassReturnDate(new Date().toISOString().split('T')[0]);
      setPhonePassStep(1);
    } catch (err) {
      setPhonePasses(prev => [{ ...passData, status: 'ISSUED', isLate: false, createdAt: now.toISOString() }, ...prev]);
      setShowIssuePhonePassModal(false);
      setPhonePassSelectedStudent(null);
      setPhonePassReturnDate(new Date().toISOString().split('T')[0]);
      setPhonePassStep(1);
    }
  };

  const handlePickupPhonePass = async (passId) => {
    try {
      const res = await fetch(`/api/phone-passes/${passId}/pickup`, {
        method: 'PUT'
      });
      if (res.ok) {
        const updatedPass = await res.json();
        setPhonePasses(prev => prev.map(p => p.id === passId ? updatedPass : p));
      } else {
        const now = new Date();
        setPhonePasses(prev => prev.map(p => {
          if (p.id === passId) {
            const oldStart = new Date(p.startTime);
            const oldAllowed = new Date(p.allowedUntil);
            const durationMs = oldAllowed.getTime() - oldStart.getTime();
            const newAllowed = new Date(now.getTime() + (durationMs > 0 ? durationMs : 3600000));
            return { ...p, status: 'OUT', startTime: now.toISOString(), allowedUntil: newAllowed.toISOString() };
          }
          return p;
        }));
      }
    } catch (err) {
      const now = new Date();
      setPhonePasses(prev => prev.map(p => {
        if (p.id === passId) {
          const oldStart = new Date(p.startTime);
          const oldAllowed = new Date(p.allowedUntil);
          const durationMs = oldAllowed.getTime() - oldStart.getTime();
          const newAllowed = new Date(now.getTime() + (durationMs > 0 ? durationMs : 3600000));
          return { ...p, status: 'OUT', startTime: now.toISOString(), allowedUntil: newAllowed.toISOString() };
        }
        return p;
      }));
    }
  };

  const handleReturnPhonePass = async (passId) => {
    try {
      const res = await fetch(`/api/phone-passes/${passId}/return`, {
        method: 'PUT'
      });
      if (res.ok) {
        const updatedPass = await res.json();
        setPhonePasses(prev => prev.map(p => p.id === passId ? updatedPass : p));
      } else {
        const now = new Date();
        setPhonePasses(prev => prev.map(p => {
          if (p.id === passId) {
            const isLate = now > new Date(p.allowedUntil);
            return { ...p, status: 'IN', returnTime: now.toISOString(), isLate };
          }
          return p;
        }));
      }
    } catch (err) {
      const now = new Date();
      setPhonePasses(prev => prev.map(p => {
        if (p.id === passId) {
          const isLate = now > new Date(p.allowedUntil);
          return { ...p, status: 'IN', returnTime: now.toISOString(), isLate };
        }
        return p;
      }));
    }
  };

  const handleDeletePhonePass = async (passId, studentName) => {
    if (!confirm(`Are you sure you want to delete the phone pass for ${studentName}? This action cannot be undone.`)) {
      return;
    }
    try {
      const res = await fetch(`/api/phone-passes/${passId}`, {
        method: 'DELETE'
      });
      if (res.ok) {
        setPhonePasses(prev => prev.filter(p => p.id !== passId));
      } else {
        alert("Failed to delete phone pass.");
      }
    } catch (err) {
      console.error("Error deleting phone pass:", err);
      setPhonePasses(prev => prev.filter(p => p.id !== passId));
    }
  };

  const handleClearAllPhonePasses = async () => {
    if (!confirm("⚠️ Are you sure you want to DELETE ALL PHONE PASSES? This will clear all pass records for all students.")) {
      return;
    }
    // Instantly clear local state for instant UI update
    setPhonePasses([]);
    localStorage.removeItem('student_tracker_phone_passes');

    try {
      const res = await fetch('/api/phone-passes/all', {
        method: 'DELETE'
      });
      if (res.ok) {
        alert("Successfully cleared all phone passes!");
      } else {
        alert("Cleared local phone passes. (Server response error)");
      }
    } catch (err) {
      console.error("Error clearing all phone passes:", err);
      alert("Cleared local phone passes.");
    }
  };

  const handleUpdateStudentPhoneDetails = async (studentId, phoneType, phoneModel, registerNumber) => {
    try {
      const res = await fetch(`/api/students/${studentId}/phone-type`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          phone_type: phoneType,
          phone_model: phoneModel,
          register_number: registerNumber
        })
      });
      if (res.ok) {
        const updatedStudent = await res.json();
        setStudents(prev => prev.map(s => s.id === studentId ? { ...s, phoneType: updatedStudent.phoneType, phoneModel: updatedStudent.phoneModel, registerNumber: updatedStudent.registerNumber } : s));
      } else {
        setStudents(prev => prev.map(s => s.id === studentId ? { ...s, phoneType, phoneModel, registerNumber } : s));
      }
    } catch (err) {
      setStudents(prev => prev.map(s => s.id === studentId ? { ...s, phoneType, phoneModel, registerNumber } : s));
    }
  };

  const ALL_PERMISSIONS = [
    { key: 'scoring', label: 'Scoring' },
    { key: 'attendance', label: 'Attendance' },
    { key: 'morning_bliss', label: 'Morning Bliss' },
    { key: 'performance', label: 'Performance' },
    { key: 'mentor', label: 'Mentor' },
    { key: 'user_sheet', label: 'User Sheet' },
    { key: 'admin_sheet', label: 'Admin Sheet' },
    { key: 'view_report', label: 'View Reports' },
    { key: 'spot_fine', label: 'Spot Fine' },
    { key: 'ineligible', label: 'Ineligible' },
    { key: 'phone_pass', label: 'Phone Pass View' },
    { key: 'phone_pass_issue', label: 'Issue Phone Pass' },
    { key: 'admin_phone_pass', label: 'Admin Phone Pass' },
  ];

  const fetchUsers = async () => {
    setUsersLoading(true);
    try {
      const res = await fetch('/api/users', { headers: getAuthHeader() });
      if (res.ok) setUsersList(await res.json());
    } catch { } finally { setUsersLoading(false); }
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
      try { return JSON.parse(saved); } catch (e) { }
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
      try { return JSON.parse(saved); } catch (e) { }
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
  const [selectedIRStudentIds, setSelectedIRStudentIds] = useState([]);
  const [irStudentSearch, setIrStudentSearch] = useState('');
  const [irAssignedStudents, setIrAssignedStudents] = useState([]);
  const [irHistoryLogs, setIrHistoryLogs] = useState([]);

  useEffect(() => {
    if (showIRModal && irAssignedStudents.length > 0) {
      const fetchHistory = async () => {
        try {
          const res = await fetch('/api/history');
          if (res.ok) {
            const data = await res.json();
            setIrHistoryLogs(data);
          }
        } catch (err) {
          console.error("Failed to fetch history:", err);
        }
      };
      fetchHistory();
    }
  }, [showIRModal, irAssignedStudents]);

  // Activities Report modal states
  const [showActivitiesReportModal, setShowActivitiesReportModal] = useState(false);
  const [activitiesLogs, setActivitiesLogs] = useState([]);
  const [activityDateFilter, setActivityDateFilter] = useState('');
  const [activityClassFilter, setActivityClassFilter] = useState('ALL');
  const [activitySearchQuery, setActivitySearchQuery] = useState('');
  const [activityEditModal, setActivityEditModal] = useState(null);
  const [showAddActivityModal, setShowAddActivityModal] = useState(false);
  const [selectedActivityIds, setSelectedActivityIds] = useState([]);
  const [newActivityForm, setNewActivityForm] = useState({
    student_id: '',
    event_type: 'tally',
    amount: 1,
    reason: '',
    date: new Date().toISOString().split('T')[0]
  });


  const fetchAllActivities = async () => {
    try {
      const res = await fetch('/api/history');
      if (res.ok) {
        const data = await res.json();
        setActivitiesLogs(data);
      }
    } catch (err) {
      console.error('Error fetching activity history:', err);
    }
  };

  useEffect(() => {
    if (showActivitiesReportModal) {
      fetchAllActivities();
    }
  }, [showActivitiesReportModal]);

  const handleSaveActivityEdit = async (e) => {
    e.preventDefault();
    if (!activityEditModal) return;
    try {
      const res = await fetch(`/api/history/${activityEditModal.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(activityEditModal)
      });
      if (res.ok) {
        const updated = await res.json();
        setActivitiesLogs(prev => prev.map(item => String(item.id) === String(updated.id) ? updated : item));
        setActivityEditModal(null);
      }
    } catch (err) {
      console.error('Failed to update activity:', err);
    }
  };

  const handleDeleteActivity = async (id) => {
    if (!window.confirm('Are you sure you want to delete this activity record?')) return;
    try {
      const res = await fetch(`/api/history/${id}`, { method: 'DELETE' });
      if (res.ok) {
        setActivitiesLogs(prev => prev.filter(item => String(item.id) !== String(id)));
        setSelectedActivityIds(prev => prev.filter(item => String(item) !== String(id)));
      }
    } catch (err) {
      console.error('Failed to delete activity:', err);
    }
  };

  const handleBulkDeleteActivities = async () => {
    if (selectedActivityIds.length === 0) return;
    if (!window.confirm(`Are you sure you want to delete ${selectedActivityIds.length} selected activity record(s)?`)) return;

    for (const id of selectedActivityIds) {
      try {
        await fetch(`/api/history/${id}`, { method: 'DELETE' });
      } catch (err) {
        console.error('Failed to delete activity:', err);
      }
    }
    setActivitiesLogs(prev => prev.filter(item => !selectedActivityIds.includes(item.id)));
    setSelectedActivityIds([]);
  };


  const handleCreateActivity = async (e) => {
    e.preventDefault();
    if (!newActivityForm.student_id || !newActivityForm.event_type) return;
    try {
      const res = await fetch('/api/history', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(newActivityForm)
      });
      if (res.ok) {
        const created = await res.json();
        setActivitiesLogs(prev => [created, ...prev]);
        setShowAddActivityModal(false);
        setNewActivityForm({ student_id: '', event_type: 'tally', amount: 1, reason: '', date: new Date().toISOString().split('T')[0] });
      }
    } catch (err) {
      console.error('Failed to create activity:', err);
    }
  };

  const filteredActivities = activitiesLogs.filter(log => {
    // 1. Date Filter
    if (activityDateFilter) {
      const logDate = log.date ? new Date(log.date).toISOString().split('T')[0] : '';
      if (logDate !== activityDateFilter) return false;
    }

    // 2. Class Filter
    const studentObj = students.find(s =>
      String(s.id) === String(log.student_id) ||
      (log.student_name && s.name.trim().toLowerCase() === String(log.student_name).trim().toLowerCase())
    );
    if (activityClassFilter !== 'ALL') {
      const cls = studentObj ? studentObj.class : (log.student_class || log.class || '');
      if (!cls || cls.trim().toLowerCase() !== activityClassFilter.toLowerCase()) {
        return false;
      }
    }

    // 3. Search Query
    if (activitySearchQuery.trim()) {
      const q = activitySearchQuery.toLowerCase().trim();
      const sName = studentObj ? studentObj.name.toLowerCase() : (log.student_name || '').toLowerCase();
      const sReason = (log.reason || '').toLowerCase();
      const sEvent = (log.event_type || '').toLowerCase();
      if (!sName.includes(q) && !sReason.includes(q) && !sEvent.includes(q)) {
        return false;
      }
    }

    return true;
  });

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

  const getFineCount = (s) => {
    if (!s) return 0;
    if (s.fineCount !== undefined && Number(s.fineCount) > 0) return Number(s.fineCount);
    if (s.history_fine_count !== undefined && Number(s.history_fine_count) > 0) return Number(s.history_fine_count);
    if (Number(s.fine) > 0 || Number(s.spotFine) > 0) return 1;
    return 0;
  };

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
    const clsName = (selectedClass || (selectedStudents[0] ? selectedStudents[0].class : '')).toLowerCase();
    if (clsName) {
      lines.push(clsName);
    }
    const starEmoji = '⭐'.repeat(sessionStar);
    selectedStudents.forEach(student => {
      lines.push(`${student.name} ${sessionStar}star ${starEmoji}`);
    });
    if (sessionStarReason.trim()) {
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
            else if (type === 'spot fine' || type === 'room fine' || type === 'spotfine' || type === 'roomfine' || type === 'fine') dynFine += 1;
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
      <div
        onClick={() => setShowSplash(false)}
        className="fixed inset-0 flex flex-col items-center justify-center bg-gradient-to-b from-[#0F172A] via-[#1A365D] to-[#0F172A] text-white z-[9999] overflow-hidden select-none animate-splash-exit font-sans cursor-pointer"
        title="Click anywhere to continue"
      >
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

        {/* Header Right Action (User Info & Logout Button) */}
        {currentUser && (
          <div className="flex items-center gap-2">
            <div className="hidden xs:flex flex-col items-end text-right">
              <span className="text-xs font-black text-white leading-tight">{currentUser.username}</span>
              <span className="text-[9px] font-bold text-slate-300 uppercase tracking-wider">
                {currentUser.role === 'super_admin'
                  ? 'Super Admin'
                  : currentUser.role === 'admin'
                    ? 'Admin'
                    : (currentUser.permissions && (currentUser.permissions.includes('admin_sheet') || currentUser.permissions.includes('admin_phone_pass')))
                      ? 'Partial Admin'
                      : 'User'}
              </span>
            </div>
            <button
              onClick={handleLogout}
              className="p-2 px-3 bg-white/10 hover:bg-rose-600/90 text-white rounded-xl transition-all shadow-xs flex items-center gap-1.5 font-extrabold text-xs cursor-pointer border border-white/20 active:scale-95"
              title="Logout from Caliph Tracer"
            >
              <Lock className="w-3.5 h-3.5 text-rose-300" />
              <span>Logout</span>
            </button>
          </div>
        )}
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
                                className={`group flex justify-between items-center p-3.5 px-4 rounded-xl border transition-all duration-200 cursor-pointer ${isSelected
                                    ? 'border-[#1A365D] bg-[#1A365D]/5 shadow-sm'
                                    : 'border-slate-200 bg-white hover:bg-slate-50 hover:border-slate-300'
                                  }`}
                              >
                                <div className="flex items-center gap-3 flex-1 min-w-0 mr-2">
                                  <div
                                    className={`w-4 h-4 rounded border flex items-center justify-center transition-all shrink-0 ${isSelected
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
                                className={`px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition-all ${status === 'present'
                                    ? 'bg-emerald-600 text-white shadow-xs'
                                    : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                                  }`}
                              >
                                Present
                              </button>

                              <button
                                type="button"
                                onClick={() => setAttendanceMap(prev => ({ ...prev, [student.id]: 'absent' }))}
                                className={`px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition-all ${status === 'absent'
                                    ? 'bg-rose-600 text-white shadow-xs'
                                    : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                                  }`}
                              >
                                Absent
                              </button>

                              <button
                                type="button"
                                onClick={() => setAttendanceMap(prev => ({ ...prev, [student.id]: 'leave' }))}
                                className={`px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition-all ${status === 'leave'
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
              if (selectedClassMentorIndex !== null) {
                const classMentor = classMentors[selectedClassMentorIndex];
                const classStudents = students.filter(s => (s.class || '').toLowerCase() === (classMentor.classAssigned || '').toLowerCase());
                const filteredClassStudents = classStudents.filter(s =>
                  s.name.toLowerCase().includes(classMentorSearch.toLowerCase()) ||
                  (s.roomNumber || '').toLowerCase().includes(classMentorSearch.toLowerCase())
                );

                return (
                  <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                    {/* Header bar */}
                    <div className="flex flex-wrap items-center justify-between gap-3 bg-white p-3 rounded-xl border border-slate-200 shadow-sm shrink-0">
                      <div className="flex items-center gap-3">
                        <button
                          onClick={() => setSelectedClassMentorIndex(null)}
                          className="w-8 h-8 rounded-lg bg-slate-100 flex items-center justify-center text-slate-600 hover:bg-[#1A365D] hover:text-white transition-colors"
                          title="Go Back"
                        >
                          <ChevronLeft className="w-5 h-5" />
                        </button>
                        <div className="w-10 h-10 rounded-xl bg-indigo-600 text-white text-sm font-extrabold flex items-center justify-center shadow-sm shrink-0">
                          {getInitials(classMentor.name)}
                        </div>
                        <div>
                          <h3 className="text-sm font-extrabold text-[#333333] leading-tight">{classMentor.name}</h3>
                          <p className="text-[11px] text-indigo-600 font-bold mt-0.5 uppercase tracking-wide">
                            Class Mentor • Class {classMentor.classAssigned.toUpperCase()} ({classStudents.length} Students)
                          </p>
                        </div>
                      </div>

                      {/* Action buttons: Download & IR */}
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => handleDownloadClassMentorSheet(classMentor)}
                          className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-extrabold rounded-xl shadow-sm transition-all active:scale-[0.98]"
                          title="Download Excel Sheet"
                        >
                          <Download className="w-4 h-4" />
                          <span>Download</span>
                        </button>
                        <button
                          onClick={() => handleGenerateClassMentorIR(classMentor)}
                          className="flex items-center gap-1.5 px-3.5 py-2 bg-[#1A365D] hover:bg-[#2A4365] text-white text-xs font-extrabold rounded-xl shadow-sm transition-all active:scale-[0.98]"
                          title="Mentor Summary IR Report"
                        >
                          <FileText className="w-4 h-4" />
                          <span>IR Button</span>
                        </button>
                      </div>
                    </div>

                    {/* Class Sheet Content */}
                    <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex flex-col gap-3 flex-1 overflow-hidden">
                      <div className="flex items-center justify-between gap-3">
                        <div className="flex items-center gap-2">
                          <Table className="w-4 h-4 text-indigo-600" />
                          <h4 className="text-xs font-extrabold text-slate-800 uppercase tracking-wider">
                            Class {classMentor.classAssigned.toUpperCase()} Student Sheet
                          </h4>
                        </div>
                        <input
                          type="text"
                          placeholder="Search student or room..."
                          value={classMentorSearch}
                          onChange={(e) => setClassMentorSearch(e.target.value)}
                          className="px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700 focus:outline-none focus:border-indigo-500 w-48"
                        />
                      </div>

                      <div className="overflow-x-auto flex-1 border border-slate-200 rounded-xl">
                        <table className="w-full text-left text-xs border-collapse">
                          <thead>
                            <tr className="bg-slate-100/80 border-b border-slate-200 text-[10px] font-extrabold text-slate-600 uppercase tracking-wider">
                              <th className="p-3 w-12 text-center">#</th>
                              <th className="p-3 min-w-[200px]">Student Name</th>
                              <th className="p-3">Class</th>
                              <th className="p-3">Room / Hostel</th>
                              <th className="p-3 text-center">Tally</th>
                              <th className="p-3 text-right">Fine (₹)</th>
                              <th className="p-3 text-center">Status</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-medium">
                            {filteredClassStudents.length > 0 ? (
                              filteredClassStudents.map((st, idx) => (
                                <tr key={st.id} className="hover:bg-slate-50/80 transition-colors">
                                  <td className="p-3 text-center text-slate-400 font-bold">{idx + 1}</td>
                                  <td className="p-3 font-extrabold text-slate-800 min-w-[200px] whitespace-nowrap">{st.name}</td>
                                  <td className="p-3 uppercase font-bold text-indigo-600">{st.class}</td>
                                  <td className="p-3 text-slate-500 font-semibold">{st.roomNumber || 'N/A'} {st.hostelBlock ? `(${st.hostelBlock})` : ''}</td>
                                  <td className="p-3 text-center">
                                    <span className="px-2 py-0.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-md font-bold text-[11px]">
                                      {st.tally || 0}
                                    </span>
                                  </td>
                                  <td className="p-3 text-right font-extrabold text-rose-600">₹{st.fine || 0}</td>
                                  <td className="p-3 text-center">
                                    {st.ineligible ? (
                                      <span className="px-2 py-0.5 bg-rose-100 text-rose-700 rounded-md font-bold text-[10px]" title={st.ineligibleReason}>
                                        Ineligible
                                      </span>
                                    ) : (
                                      <span className="px-2 py-0.5 bg-emerald-100 text-emerald-700 rounded-md font-bold text-[10px]">
                                        Active
                                      </span>
                                    )}
                                  </td>
                                </tr>
                              ))
                            ) : (
                              <tr>
                                <td colSpan="7" className="p-8 text-center text-slate-400 text-xs">
                                  No students found for Class {classMentor.classAssigned.toUpperCase()}.
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                );
              }

              if (selectedMentorIndex === null) {
                return (
                  <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                    {/* Top Switcher: Room Mentors vs Class Mentors */}
                    <div className="flex bg-slate-200/60 p-1 rounded-xl shrink-0">
                      <button
                        onClick={() => setMentorTypeTab('room')}
                        className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${mentorTypeTab === 'room'
                            ? 'bg-white text-[#1A365D] shadow-sm'
                            : 'text-slate-500 hover:text-slate-700'
                          }`}
                      >
                        👥 Room Mentors ({mentors.length})
                      </button>
                      <button
                        onClick={() => setMentorTypeTab('class')}
                        className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${mentorTypeTab === 'class'
                            ? 'bg-white text-[#1A365D] shadow-sm'
                            : 'text-slate-500 hover:text-slate-700'
                          }`}
                      >
                        🏫 Class Mentors ({classMentors.length})
                      </button>
                    </div>

                    {mentorTypeTab === 'room' ? (
                      /* All Mentors Quick Directory */
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
                    ) : (
                      /* Class Mentors View */
                      <div className="flex flex-col gap-3">
                        <div className="flex items-center justify-between mb-4 mt-2">
                          <span className="text-sm font-extrabold text-[#1A365D] uppercase tracking-wider flex items-center gap-2">
                            <School className="w-5 h-5 text-[#1A365D]" />
                            All Class Mentors ({classMentors.length})
                          </span>
                          <div className="flex items-center gap-2">
                            <button
                              onClick={() => setShowAddClassMentorModal(true)}
                              className="flex items-center gap-1.5 px-4 py-2 bg-[#1A365D] hover:bg-[#2A4365] text-white text-xs font-extrabold rounded-xl shadow-md hover:shadow-lg transition-all active:scale-[0.98]"
                            >
                              <Plus className="w-4 h-4" />
                              <span>Add Class Mentor</span>
                            </button>
                            <button
                              onClick={() => setShowDeleteClassMentorModal(true)}
                              className="flex items-center gap-1.5 px-4 py-2 bg-rose-50 border border-rose-200 hover:bg-rose-600 hover:text-white hover:border-rose-600 text-rose-600 text-xs font-extrabold rounded-xl shadow-sm hover:shadow-md transition-all active:scale-[0.98]"
                            >
                              <Trash2 className="w-4 h-4" />
                              <span>Remove</span>
                            </button>
                          </div>
                        </div>

                        {classMentors.length === 0 ? (
                          <div className="text-center py-12 flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-2xl bg-white p-6">
                            <School className="w-12 h-12 text-slate-300 mb-3" />
                            <h4 className="text-sm font-extrabold text-slate-700">No Class Mentors Created</h4>
                            <p className="text-xs text-slate-400 mt-1 mb-4">Click "Add Class Mentor" above to create a folder for a class mentor.</p>
                            <button
                              onClick={() => setShowAddClassMentorModal(true)}
                              className="px-4 py-2 bg-[#1A365D] text-white text-xs font-extrabold rounded-xl shadow-sm hover:bg-[#2A4365] transition-colors flex items-center gap-2"
                            >
                              <Plus className="w-4 h-4" />
                              <span>Add Class Mentor</span>
                            </button>
                          </div>
                        ) : (
                          <div className="grid grid-cols-1 gap-3">
                            {classMentors.map((cm, cmIdx) => {
                              const classCount = students.filter(s => (s.class || '').toLowerCase() === (cm.classAssigned || '').toLowerCase()).length;
                              return (
                                <button
                                  key={cm.id}
                                  onClick={() => setSelectedClassMentorIndex(cmIdx)}
                                  className="group relative bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex items-center justify-between transition-all duration-200 hover:shadow-md hover:border-indigo-300 active:scale-[0.98] overflow-hidden text-left"
                                >
                                  <div className="flex items-center gap-4">
                                    <div className="w-12 h-12 rounded-xl text-lg font-extrabold flex items-center justify-center transition-all duration-300 bg-indigo-50 text-indigo-700 group-hover:scale-110 group-hover:bg-indigo-100 shadow-inner">
                                      {getInitials(cm.name)}
                                    </div>
                                    <div className="text-left flex flex-col justify-center">
                                      <h4 className="text-[15px] font-extrabold text-slate-800 tracking-tight leading-tight">{cm.name}</h4>
                                      <span className="text-xs text-indigo-600 font-extrabold mt-0.5 uppercase tracking-wide">Class {cm.classAssigned.toUpperCase()}</span>
                                    </div>
                                  </div>
                                  <div className="flex items-center gap-3">
                                    <span className="text-[10px] font-extrabold text-slate-500 bg-slate-100 px-2.5 py-1 rounded-md">
                                      {classCount} Students
                                    </span>
                                    <ChevronRight className="w-5 h-5 text-indigo-500 group-hover:translate-x-0.5 transition-all duration-300" />
                                  </div>
                                </button>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    )}
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
                      className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${mentorSubTab === 'diary'
                          ? 'bg-white text-[#1A365D] shadow-sm'
                          : 'text-slate-500 hover:text-slate-700'
                        }`}
                    >
                      📖 Diary
                    </button>
                    <button
                      onClick={() => setMentorSubTab('attendance')}
                      className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${mentorSubTab === 'attendance'
                          ? 'bg-white text-[#1A365D] shadow-sm'
                          : 'text-slate-500 hover:text-slate-700'
                        }`}
                    >
                      🌙 Night Attendance
                    </button>
                    <button
                      onClick={() => setMentorSubTab('summary')}
                      className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${mentorSubTab === 'summary'
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
                                    className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${isSelected
                                        ? 'border-[#1A365D] bg-[#1A365D]/5 ring-1 ring-[#1A365D]/20 shadow-sm'
                                        : 'border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50'
                                      }`}
                                  >
                                    <div className="flex items-center gap-3">
                                      <div className={`w-9 h-9 rounded-full flex items-center justify-center text-[11px] font-extrabold transition-colors ${isSelected ? 'bg-[#1A365D] text-white' : 'bg-slate-100 text-[#1A365D]'
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
                                  className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${isSelected
                                      ? 'border-amber-500 bg-amber-500/5 ring-1 ring-amber-500/20 shadow-sm'
                                      : 'border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50'
                                    }`}
                                >
                                  <div className="flex items-center gap-3">
                                    <div className={`w-9 h-9 rounded-full flex items-center justify-center text-[11px] font-extrabold transition-colors ${isSelected ? 'bg-amber-500 text-white' : 'bg-slate-100 text-amber-600'
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
                                    if (!ds) return '';
                                    const p = ds.split('-');
                                    if (p.length === 3) {
                                      const day = parseInt(p[2], 10);
                                      const month = parseInt(p[1], 10);
                                      const year = p[0].slice(-2);
                                      return `${day}/${month}/${year}`;
                                    }
                                    return ds;
                                  };
                                  const start = formatStr(diaryStartDate);
                                  const end = formatStr(diaryEndDate);

                                  let rawText = `Date from ${start}   to   ${end}\nMentor: ${mentor.name}\n`;

                                  // Group assigned students by class
                                  const groupedByClass = {};
                                  assignedStudents.forEach(s => {
                                    const cls = (s.class || mentor.classAssigned || 'GENERAL').toUpperCase();
                                    if (!groupedByClass[cls]) {
                                      groupedByClass[cls] = [];
                                    }
                                    groupedByClass[cls].push(s);
                                  });

                                  Object.keys(groupedByClass).forEach(cls => {
                                    rawText += `\n${cls}\n`;
                                    groupedByClass[cls].forEach(s => {
                                      const record = diaryRecords[s.id] || { status: 'none', days: '' };
                                      if (record.status === 'written') {
                                        rawText += `${s.name}✅\n`;
                                      } else if (record.status === 'not_written') {
                                        const daysStr = (record.days || '').replace(/[^0-9]/g, '');
                                        const numDays = parseInt(daysStr) || 0;
                                        const tallies = (numDays * (numDays + 1)) / 2;
                                        rawText += `${s.name} (not submitted ${record.days || numDays + ' day'}) ${tallies} tally\n`;
                                      } else {
                                        rawText += `${s.name} (not submitted)\n`;
                                      }
                                    });
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
                                            className={`w-full text-xs py-1.5 px-3 rounded-lg focus:outline-none focus:ring-2 font-medium transition-colors ${currentStatus === 'not_written'
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
                                  const timeStr = today.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

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
                                setSelectedIRStudentIds(assignedStudents.map(s => s.id));
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
                          </div>
                        </div>


                        <div className="w-full flex-1 flex flex-col gap-2 mx-auto overflow-hidden bg-white border border-slate-200 rounded-xl shadow-sm">
                          <div className="overflow-x-auto h-full">
                            <table className="w-full text-left text-xs">
                              <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider sticky top-0 border-b border-slate-200 z-10 shadow-sm">
                                <tr>
                                  <th className="p-3 font-extrabold text-center">Class</th>
                                  <th className="p-3 font-extrabold text-center min-w-[200px]">Name</th>
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
                                  );
                                }) : (
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
                            onClick={() => {
                              if (isStopwatchRunning) {
                                setIsStopwatchRunning(false);
                                stopwatchAccumulatedRef.current = morningBlissDuration;
                              } else {
                                stopwatchStartTimeRef.current = Date.now() - stopwatchAccumulatedRef.current;
                                setIsStopwatchRunning(true);
                              }
                            }}
                            className={`px-4 py-2 rounded-lg font-bold text-white transition-colors ${isStopwatchRunning ? 'bg-red-500 hover:bg-red-600' : 'bg-emerald-500 hover:bg-emerald-600'}`}
                          >
                            {isStopwatchRunning ? 'Stop' : 'Start'}
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setIsStopwatchRunning(false);
                              stopwatchAccumulatedRef.current = 0;
                              stopwatchStartTimeRef.current = null;
                              setMorningBlissDuration(0);
                            }}
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
                    {(() => {
                      const todayDateStr = new Date().toISOString().split('T')[0];
                      const visibleMbStudents = students.filter(s => {
                        const hasMbData = (s.morningBlissMark != null && s.morningBlissMark !== '') || s.morningBlissScript != null || (s.morningBlissTopic && s.morningBlissTopic.trim() !== '');
                        if (!hasMbData) return false;

                        const studentDate = s.summaryDate || todayDateStr;
                        if (mbFromDate && studentDate < mbFromDate) return false;
                        if (mbToDate && studentDate > mbToDate) return false;
                        return true;
                      });

                      const allVisibleSelected = visibleMbStudents.length > 0 && visibleMbStudents.every(s => selectedMbStudentIds.includes(s.id));

                      return (
                        <>
                          <div className="p-4 bg-white border-b border-slate-200 shrink-0 flex flex-wrap justify-between items-center gap-3">
                            <div className="flex flex-col">
                              <span className="font-bold text-slate-800">Today's Morning Bliss Results</span>
                              <span className="text-xs font-medium text-slate-500">Summary Date: {todayDateStr}</span>
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

                              {/* Delete Selected Button */}
                              {selectedMbStudentIds.length > 0 && (
                                <button
                                  onClick={() => handleDeleteSelectedMorningBliss()}
                                  className="bg-red-600 hover:bg-red-700 text-white px-3 py-2 rounded-lg font-bold text-sm transition-all shadow-sm flex items-center gap-1.5 animate-fade-in"
                                  title="Delete selected Morning Bliss records"
                                >
                                  <Trash2 className="w-4 h-4" /> Delete Selected ({selectedMbStudentIds.length})
                                </button>
                              )}

                              <button
                                onClick={() => {
                                  if (window.confirm("Are you sure you want to clear all Morning Bliss data for today?")) {
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
                                    setSelectedMbStudentIds([]);

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
                                CLEAR ALL
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
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider w-10 text-center">
                                        <input
                                          type="checkbox"
                                          checked={allVisibleSelected}
                                          onChange={(e) => {
                                            if (e.target.checked) {
                                              const visibleIds = visibleMbStudents.map(s => s.id);
                                              setSelectedMbStudentIds(Array.from(new Set([...selectedMbStudentIds, ...visibleIds])));
                                            } else {
                                              const visibleIds = visibleMbStudents.map(s => s.id);
                                              setSelectedMbStudentIds(selectedMbStudentIds.filter(id => !visibleIds.includes(id)));
                                            }
                                          }}
                                          className="w-4 h-4 accent-red-600 rounded cursor-pointer"
                                          title="Select / Deselect all visible records"
                                        />
                                      </th>
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Class</th>
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Name</th>
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Mark</th>
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Script Mark</th>
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Total</th>
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">Star</th>
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider">EV</th>
                                      <th className="p-3 text-xs font-extrabold text-slate-500 uppercase tracking-wider text-center">Action</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {(() => {
                                      const rows = CLASSES.flatMap(c => {
                                        const classStudents = visibleMbStudents.filter(s => s.class === c);
                                        if (classStudents.length === 0) return [];

                                        return classStudents.map(s => {
                                          const total = (Number(s.morningBlissMark) || 0) + (Number(s.morningBlissScript) || 0);
                                          const isSelected = selectedMbStudentIds.includes(s.id);
                                          return (
                                            <tr key={s.id} className={`border-b border-slate-100 transition-colors ${isSelected ? 'bg-red-50/50' : 'hover:bg-slate-50'}`}>
                                              <td className="p-3 text-center">
                                                <input
                                                  type="checkbox"
                                                  checked={isSelected}
                                                  onChange={(e) => {
                                                    if (e.target.checked) {
                                                      setSelectedMbStudentIds([...selectedMbStudentIds, s.id]);
                                                    } else {
                                                      setSelectedMbStudentIds(selectedMbStudentIds.filter(id => id !== s.id));
                                                    }
                                                  }}
                                                  className="w-4 h-4 accent-red-600 rounded cursor-pointer"
                                                />
                                              </td>
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
                                              <td className="p-3 text-center">
                                                <button
                                                  onClick={() => handleDeleteSelectedMorningBliss([s.id])}
                                                  title={`Delete ${s.name}'s Morning Bliss result`}
                                                  className="p-1.5 text-red-500 hover:text-red-700 hover:bg-red-100/60 rounded-lg transition-colors inline-flex items-center justify-center"
                                                >
                                                  <Trash2 className="w-4 h-4" />
                                                </button>
                                              </td>
                                            </tr>
                                          )
                                        });
                                      });

                                      return rows.length === 0 ? (
                                        <tr>
                                          <td colSpan="9" className="p-8 text-center text-slate-400 font-medium">
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
                        </>
                      );
                    })()}
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
                          onChange={(e) => {
                            setSheetClass(e.target.value);
                            setSheetSelectedStudentIds([]);
                            setShowSheetDropdown(true);
                          }}
                          className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50 appearance-none cursor-pointer"
                          required
                        >
                          <option value="" disabled>Choose a class...</option>
                          <option value="ALL">🌟 ALL CLASSES ({students.length} Students)</option>
                          {visibleClasses.map(cls => <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>)}
                        </select>
                      </div>

                      {/* Multi-selected Student Chips */}
                      {sheetSelectedStudentIds.length > 0 && (
                        <div className="flex flex-wrap gap-1.5 p-3 bg-blue-50/80 border border-blue-200 rounded-xl">
                          <div className="w-full flex items-center justify-between text-xs font-extrabold text-blue-900 mb-1">
                            <span>Selected Students ({sheetSelectedStudentIds.length}):</span>
                            <button
                              type="button"
                              onClick={() => setSheetSelectedStudentIds([])}
                              className="text-rose-600 hover:underline text-[11px] font-bold cursor-pointer"
                            >
                              Clear All
                            </button>
                          </div>
                          {sheetSelectedStudentIds.map(id => {
                            const st = students.find(s => s.id === id);
                            if (!st) return null;
                            return (
                              <span
                                key={id}
                                className="inline-flex items-center gap-1.5 px-3 py-1 bg-white border border-blue-200 text-blue-900 rounded-lg text-xs font-bold shadow-2xs"
                              >
                                {st.name} <span className="text-[10px] text-slate-400 font-extrabold">({st.class.toUpperCase()})</span>
                                <button
                                  type="button"
                                  onClick={() => setSheetSelectedStudentIds(prev => prev.filter(item => item !== id))}
                                  className="text-slate-400 hover:text-rose-600 font-bold ml-1 cursor-pointer"
                                >
                                  ✕
                                </button>
                              </span>
                            );
                          })}
                        </div>
                      )}

                      <div className="relative">
                        <div className="flex items-center justify-between mb-2">
                          <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider">
                            Student Name(s) {sheetSelectedStudentIds.length > 0 ? `(${sheetSelectedStudentIds.length} Selected)` : ''}
                          </label>
                          {sheetClass && (
                            <button
                              type="button"
                              onClick={() => setShowSheetDropdown(!showSheetDropdown)}
                              className="text-xs font-extrabold text-[#1A365D] hover:underline flex items-center gap-1 cursor-pointer"
                            >
                              {showSheetDropdown ? 'Hide List ✕' : 'Show List ▾'}
                            </button>
                          )}
                        </div>

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
                            placeholder={sheetSelectedStudentIds.length > 0 ? "Search more students..." : "Type student name to search..."}
                            className="w-full p-4 pl-12 pr-10 text-base border border-slate-200 rounded-xl font-bold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                            required={sheetSelectedStudentIds.length === 0}
                          />
                          {sheetNameSearch && (
                            <button
                              type="button"
                              onClick={() => setSheetNameSearch('')}
                              className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 font-bold text-sm cursor-pointer p-1"
                            >
                              ✕
                            </button>
                          )}
                        </div>

                        {sheetClass && showSheetDropdown && (() => {
                          const availableStudents = students.filter(s => {
                            const matchesClass = sheetClass === 'ALL' ? true : s.class === sheetClass;
                            const matchesSearch = !sheetNameSearch || s.name.toLowerCase().includes(sheetNameSearch.toLowerCase()) || s.class.toLowerCase().includes(sheetNameSearch.toLowerCase());
                            return matchesClass && matchesSearch;
                          });

                          return (
                            <div className="mt-2 bg-white border border-slate-200 rounded-2xl shadow-xl max-h-64 overflow-y-auto z-30 absolute left-0 right-0 p-2 flex flex-col gap-1.5">
                              <div className="p-2.5 bg-slate-50 border-b border-slate-100 rounded-xl text-xs font-bold text-slate-600 flex items-center justify-between gap-2 shrink-0 sticky top-0 z-10">
                                <span>{availableStudents.length} Students found</span>
                                <div className="flex items-center gap-2">
                                  {availableStudents.length > 0 && (
                                    <button
                                      type="button"
                                      onClick={() => {
                                        const allFilteredIds = availableStudents.map(s => s.id);
                                        const allSelected = allFilteredIds.every(id => sheetSelectedStudentIds.includes(id));
                                        if (allSelected) {
                                          setSheetSelectedStudentIds(prev => prev.filter(id => !allFilteredIds.includes(id)));
                                        } else {
                                          setSheetSelectedStudentIds(prev => Array.from(new Set([...prev, ...allFilteredIds])));
                                        }
                                      }}
                                      className="text-blue-600 hover:underline font-black cursor-pointer text-xs"
                                    >
                                      {availableStudents.every(s => sheetSelectedStudentIds.includes(s.id)) ? 'Deselect All' : 'Select All Filtered'}
                                    </button>
                                  )}
                                  <button
                                    type="button"
                                    onClick={() => setShowSheetDropdown(false)}
                                    className="text-xs font-black bg-slate-200 hover:bg-slate-300 text-slate-700 px-2.5 py-1 rounded-lg transition-colors cursor-pointer"
                                  >
                                    Done / Hide List ✕
                                  </button>
                                </div>
                              </div>

                              {availableStudents.map(s => {
                                const isSelected = sheetSelectedStudentIds.includes(s.id);
                                return (
                                  <div
                                    key={s.id}
                                    className={`p-3 rounded-xl cursor-pointer font-bold text-xs flex items-center justify-between transition-colors border ${isSelected
                                        ? 'bg-blue-50 border-blue-300 text-blue-900 shadow-2xs'
                                        : 'hover:bg-slate-50 border-slate-100 text-slate-700'
                                      }`}
                                    onClick={() => {
                                      setSheetSelectedStudentIds(prev =>
                                        prev.includes(s.id) ? prev.filter(id => id !== s.id) : [...prev, s.id]
                                      );
                                    }}
                                  >
                                    <div className="flex items-center gap-2.5">
                                      <input
                                        type="checkbox"
                                        checked={isSelected}
                                        onChange={() => { }}
                                        className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500 cursor-pointer"
                                      />
                                      <span>{s.name}</span>
                                    </div>
                                    <span className="text-[10px] font-extrabold uppercase px-2 py-0.5 bg-slate-100 text-slate-600 rounded-md border border-slate-200">
                                      Class {s.class.toUpperCase()}
                                    </span>
                                  </div>
                                );
                              })}

                              {availableStudents.length === 0 && (
                                <div className="p-4 text-center text-slate-400 font-semibold text-sm">
                                  No matching student found {sheetClass !== 'ALL' ? `in Class ${sheetClass.toUpperCase()}` : ''}
                                </div>
                              )}
                            </div>
                          );
                        })()}
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-2 uppercase tracking-wider">Reason</label>
                        <input
                          type="text"
                          value={sheetReason}
                          onFocus={() => setShowSheetDropdown(false)}
                          onChange={(e) => setSheetReason(e.target.value)}
                          placeholder="E.g. Disruption, late coming..."
                          className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                          required
                        />
                      </div>
                      <button type="submit" className="w-full py-4 mt-2 bg-blue-600 hover:bg-blue-700 text-white font-extrabold rounded-xl transition-all shadow-md active:scale-[0.98] flex justify-center items-center gap-2 cursor-pointer">
                        <Send className="w-5 h-5" />
                        <span>
                          Submit Sheet {sheetSelectedStudentIds.length > 0 ? `(${sheetSelectedStudentIds.length} Students)` : ''}
                        </span>
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
                          onChange={(e) => {
                            setSpotClass(e.target.value);
                            setShowSpotDropdown(true);
                          }}
                          className="w-full p-4 border border-slate-200 rounded-xl text-sm font-semibold focus:outline-none focus:border-[#1A365D] bg-slate-50 appearance-none cursor-pointer"
                          required
                        >
                          <option value="" disabled>Choose a class...</option>
                          <option value="ALL">🌟 All Classes (Entire School)</option>
                          {visibleClasses.map(cls => <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>)}
                        </select>
                      </div>

                      {spotSelectedStudentIds.length > 0 && (
                        <div className="p-3 bg-rose-50/70 border border-rose-200 rounded-xl flex flex-col gap-2">
                          <div className="flex items-center justify-between text-xs font-extrabold text-rose-800">
                            <span>Selected Students ({spotSelectedStudentIds.length}):</span>
                            <button
                              type="button"
                              onClick={() => setSpotSelectedStudentIds([])}
                              className="text-[10px] text-rose-600 hover:underline font-extrabold cursor-pointer"
                            >
                              Clear All
                            </button>
                          </div>
                          <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto pr-1">
                            {spotSelectedStudentIds.map(id => {
                              const st = students.find(s => s.id === id);
                              if (!st) return null;
                              return (
                                <span
                                  key={id}
                                  className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-white border border-rose-200 text-rose-900 rounded-lg text-xs font-bold shadow-2xs"
                                >
                                  <span>{st.name} <span className="text-[10px] text-rose-600 font-extrabold">({st.class.toUpperCase()})</span></span>
                                  <button
                                    type="button"
                                    onClick={() => setSpotSelectedStudentIds(prev => prev.filter(i => i !== id))}
                                    className="hover:text-rose-600 text-slate-400 font-black ml-0.5 cursor-pointer"
                                  >
                                    ✕
                                  </button>
                                </span>
                              );
                            })}
                          </div>
                        </div>
                      )}

                      <div>
                        <div className="flex items-center justify-between mb-2">
                          <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider">Student Name(s) (Multi-Select)</label>
                          {spotClass && (
                            <button
                              type="button"
                              onClick={() => setShowSpotDropdown(!showSpotDropdown)}
                              className="text-xs font-extrabold text-[#1A365D] hover:underline flex items-center gap-1 cursor-pointer"
                            >
                              {showSpotDropdown ? 'Hide List ✕' : 'Show List ▾'}
                            </button>
                          )}
                        </div>
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
                            placeholder={spotSelectedStudentIds.length === 0 ? "Type student name to search & select..." : "Search & add more students..."}
                            className="w-full p-4 pl-12 text-sm sm:text-base border border-slate-200 rounded-xl font-bold focus:outline-none focus:border-[#1A365D] bg-slate-50"
                          />
                        </div>
                        {spotClass && showSpotDropdown && (
                          (() => {
                            const filtered = students.filter(s => {
                              const classMatch = spotClass === 'ALL' ? true : s.class === spotClass;
                              const nameMatch = !spotNameSearch.trim() || s.name.toLowerCase().includes(spotNameSearch.trim().toLowerCase());
                              return classMatch && nameMatch;
                            });

                            return (
                              <div className="mt-2 bg-white border border-slate-200 rounded-xl shadow-sm max-h-60 overflow-y-auto divide-y divide-slate-100 relative w-full">
                                {/* Action Bar */}
                                <div className="p-2.5 bg-slate-50 border-b border-slate-200 sticky top-0 z-10 flex items-center justify-between gap-2">
                                  {filtered.length > 1 ? (
                                    <button
                                      type="button"
                                      className="text-xs font-black text-emerald-800 bg-emerald-100 hover:bg-emerald-200 px-2.5 py-1 rounded-lg transition-colors cursor-pointer"
                                      onClick={() => {
                                        const matchingIds = filtered.map(s => s.id);
                                        setSpotSelectedStudentIds(prev => Array.from(new Set([...prev, ...matchingIds])));
                                      }}
                                    >
                                      Select All ({filtered.length})
                                    </button>
                                  ) : <div />}
                                  <button
                                    type="button"
                                    onClick={() => setShowSpotDropdown(false)}
                                    className="text-xs font-black bg-slate-200 hover:bg-slate-300 text-slate-700 px-2.5 py-1 rounded-lg transition-colors cursor-pointer"
                                  >
                                    Done / Hide List ✕
                                  </button>
                                </div>

                                {filtered.map(s => {
                                  const isSelected = spotSelectedStudentIds.includes(s.id);
                                  return (
                                    <div
                                      key={s.id}
                                      className={`p-3.5 hover:bg-rose-50/50 cursor-pointer font-bold text-sm text-slate-800 flex items-center justify-between transition-colors ${isSelected ? 'bg-rose-50/80 text-rose-950' : ''
                                        }`}
                                      onClick={() => {
                                        setSpotSelectedStudentIds(prev =>
                                          prev.includes(s.id) ? prev.filter(i => i !== s.id) : [...prev, s.id]
                                        );
                                      }}
                                    >
                                      <div className="flex items-center gap-2.5">
                                        <div className={`w-4 h-4 rounded border flex items-center justify-center transition-colors ${isSelected ? 'bg-rose-600 border-rose-600' : 'border-slate-300 bg-white'
                                          }`}>
                                          {isSelected && <Check className="w-3 h-3 text-white stroke-[3]" />}
                                        </div>
                                        <span>{s.name}</span>
                                      </div>
                                      <span className="text-[11px] font-extrabold px-2 py-0.5 rounded-md bg-slate-100 border border-slate-200 text-slate-700">
                                        {s.class.toUpperCase()}
                                      </span>
                                    </div>
                                  );
                                })}
                                {filtered.length === 0 && (
                                  <div className="p-4 text-center text-slate-400 font-semibold text-sm">
                                    No matching student found {spotClass !== 'ALL' ? `in Class ${spotClass.toUpperCase()}` : ''}
                                  </div>
                                )}
                              </div>
                            );
                          })()
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
                        {hostelBlocks.map(h => (
                          <div key={h.id} className="relative group/card">
                            <button
                              onClick={() => {
                                setSelectedHostel(h.name);
                                setSelectedRoom(null);
                              }}
                              className="w-full p-5 bg-white border border-slate-200 hover:border-[#1A365D] rounded-2xl shadow-sm hover:shadow-md transition-all active:scale-[0.98] flex items-center justify-between group text-left pr-12 cursor-pointer"
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
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                openEditHostelModal(h);
                              }}
                              className="absolute top-3.5 right-3.5 p-2 rounded-xl bg-slate-100 hover:bg-amber-100 text-slate-400 hover:text-amber-700 transition-all border border-slate-200 hover:border-amber-300 shadow-xs cursor-pointer"
                              title="Change Hostel Name"
                            >
                              <Pencil className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : !selectedRoom ? (
                    /* STEP 2: Show Room Buttons for Selected Hostel */
                    <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
                      {/* Header Card with Add Room, Delete Room, Rename Hostel, Switch Block */}
                      <div className="flex flex-wrap items-center justify-between bg-white p-3.5 rounded-2xl border border-slate-200 shadow-sm shrink-0 gap-2">
                        <div>
                          <span className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider block">{selectedHostel}</span>
                          <h3 className="text-[#1A365D] font-extrabold text-base">Select Room ({hostelRoomList.length} Rooms)</h3>
                        </div>
                        <div className="flex items-center gap-1.5 sm:gap-2 flex-wrap">
                          {/* 1. ADD ROOM BUTTON */}
                          <button
                            onClick={() => {
                              setNewRoomNumberInput('');
                              setShowAddNewRoomModal(true);
                            }}
                            className="text-xs font-extrabold text-emerald-800 bg-emerald-100 hover:bg-emerald-200 px-3 py-1.5 rounded-xl border border-emerald-300 transition-all flex items-center gap-1 shadow-xs active:scale-95 cursor-pointer"
                            title="Add room"
                          >
                            <Plus className="w-3.5 h-3.5" />
                            <span>Add Room</span>
                          </button>

                          {/* 2. CHANGE HOSTEL NAME BUTTON */}
                          <button
                            onClick={() => {
                              const curHostel = hostelBlocks.find(h => h.name === selectedHostel) || { id: Date.now(), name: selectedHostel };
                              openEditHostelModal(curHostel);
                            }}
                            className="text-xs font-extrabold text-amber-800 bg-amber-100 hover:bg-amber-200 px-3 py-1.5 rounded-xl border border-amber-300 transition-all flex items-center gap-1 shadow-xs active:scale-95 cursor-pointer"
                            title="Change Hostel Name"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                            <span>Change Hostel Name</span>
                          </button>

                          {/* 3. DELETE ROOM BUTTON */}
                          <button
                            onClick={() => {
                              setRoomsToDelete([]);
                              setShowDeleteRoomModal(true);
                            }}
                            className="text-xs font-extrabold text-rose-700 bg-rose-100 hover:bg-rose-200 px-3 py-1.5 rounded-xl border border-rose-200 transition-all flex items-center gap-1 shadow-xs active:scale-95 cursor-pointer"
                            title="Delete room"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            <span>Delete</span>
                          </button>

                          {/* 4. SWITCH HOSTEL BLOCK */}
                          <button
                            onClick={() => setSelectedHostel(null)}
                            className="text-xs font-bold text-slate-600 hover:text-slate-800 bg-slate-100 hover:bg-slate-200 px-3 py-1.5 rounded-xl border border-slate-200 transition-colors flex items-center gap-1 cursor-pointer"
                            title="Switch to another hostel block"
                          >
                            <ChevronLeft className="w-3.5 h-3.5" />
                            <span>Switch Block</span>
                          </button>
                        </div>
                      </div>

                      <p className="text-[11px] font-bold text-slate-400 text-center -mt-1">
                        💡 Long press any room card (or click edit icon) to rename room
                      </p>

                      {hostelRoomList.length > 0 ? (
                        <div className="grid grid-cols-3 sm:grid-cols-5 gap-3">
                          {hostelRoomList.map(num => (
                            <button
                              key={num}
                              onMouseDown={() => handleRoomPressStart(num)}
                              onMouseUp={() => handleRoomPressEnd()}
                              onMouseLeave={() => handleRoomPressEnd()}
                              onTouchStart={() => handleRoomPressStart(num)}
                              onTouchEnd={() => handleRoomPressEnd()}
                              onClick={() => {
                                if (isLongPressTriggeredRef.current) return;
                                const roomStr = `Room ${num}`;
                                setSelectedRoom(roomStr);
                                const allocated = getHostelRoomStudents(selectedHostel, num);
                                setPerformanceSelectedStudents(allocated.map(s => s.id));
                              }}
                              className="p-4 bg-white border border-slate-200 hover:border-emerald-500 hover:bg-emerald-50/40 rounded-2xl shadow-sm hover:shadow transition-all active:scale-[0.96] flex flex-col items-center justify-center gap-1.5 group relative"
                            >
                              {/* Small Edit Icon button */}
                              <div
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setEditingRoomOldNum(String(num));
                                  setEditingRoomNewNum(String(num));
                                  setShowEditRoomNumberModal(true);
                                }}
                                className="absolute top-1.5 right-1.5 p-1 rounded-md bg-slate-100 hover:bg-emerald-100 text-slate-400 hover:text-emerald-700 opacity-70 group-hover:opacity-100 transition-all"
                                title="Edit Room Number"
                              >
                                <Pencil className="w-3 h-3" />
                              </div>

                              <div className="w-9 h-9 rounded-xl bg-slate-100 group-hover:bg-emerald-600 text-slate-600 group-hover:text-white flex items-center justify-center transition-colors">
                                <School className="w-5 h-5" />
                              </div>
                              <span className="font-extrabold text-xs text-slate-800 group-hover:text-emerald-700">Room {num}</span>
                            </button>
                          ))}
                        </div>
                      ) : (
                        <div className="p-8 text-center bg-white rounded-2xl border border-slate-200 shadow-sm flex flex-col items-center justify-center gap-3 my-auto">
                          <div className="w-12 h-12 rounded-2xl bg-emerald-50 text-emerald-600 flex items-center justify-center">
                            <School className="w-6 h-6" />
                          </div>
                          <h4 className="font-extrabold text-slate-800 text-base">No Rooms in {selectedHostel}</h4>
                          <p className="text-xs font-semibold text-slate-500 max-w-xs leading-relaxed">
                            There are currently no rooms created for <strong>{selectedHostel}</strong>.
                          </p>
                          <button
                            onClick={() => {
                              setNewRoomNumberInput('');
                              setShowAddNewRoomModal(true);
                            }}
                            className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-xs rounded-xl shadow-xs transition-all active:scale-95 flex items-center gap-1.5 cursor-pointer mt-1"
                          >
                            <Plus className="w-4 h-4" /> Add Room Now
                          </button>
                        </div>
                      )}
                    </div>
                  ) : (

                    /* STEP 3: Student Selection & Room Actions inside Selected Room */
                    (() => {
                      const currentRoomKey = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
                      const allocatedRoomStudents = getHostelRoomStudents(selectedHostel, currentRoomKey);

                      const selectedCount = performanceSelectedStudents.length;


                      return (
                        <div className="flex-1 overflow-hidden flex flex-col bg-slate-50 relative">
                          {/* TOP BAR: Room Details Card at very top (replacing "Allocated Students" heading) */}
                          <div className="p-3 bg-emerald-50 border-b border-emerald-100 flex items-center justify-between shrink-0 shadow-xs flex-wrap gap-2">
                            <div className="flex items-center gap-2.5">
                              <div className="w-9 h-9 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-bold shadow-xs">
                                <School className="w-5 h-5" />
                              </div>
                              <div>
                                <span className="text-[10px] font-extrabold text-emerald-700 uppercase tracking-wider block leading-tight">{selectedHostel}</span>
                                <h3 className="text-[#1A365D] font-extrabold text-sm uppercase leading-tight">{selectedRoom} ({allocatedRoomStudents.length} Students)</h3>
                              </div>
                            </div>

                            <div className="flex items-center gap-1.5 sm:gap-2 flex-wrap">
                              {/* 1. ADD STUDENT BUTTON */}
                              <button
                                onClick={() => {
                                  setAddRoomStudentName('');
                                  setAddRoomStudentClass(CLASSES[0] || 'S1B');
                                  setAddRoomStudentModalTab('single');
                                  setShowAddRoomStudentModal(true);
                                }}
                                className="text-xs font-extrabold text-emerald-800 bg-emerald-200/90 hover:bg-emerald-300 px-2.5 sm:px-3 py-1.5 rounded-xl border border-emerald-300 transition-all flex items-center gap-1 shadow-xs active:scale-95"
                                title="Add student to this room"
                              >
                                <Plus className="w-3.5 h-3.5" />
                                <span>Add</span>
                              </button>

                              {/* 1B. BULK / EXCEL IMPORT BUTTON */}
                              <button
                                onClick={() => {
                                  setAddRoomStudentName('');
                                  setAddRoomStudentClass(CLASSES[0] || 'S1B');
                                  setAddRoomStudentModalTab('bulk');
                                  setBulkImportText('');
                                  setBulkParsedStudents([]);
                                  setBulkImportFileName('');
                                  setShowAddRoomStudentModal(true);
                                }}
                                className="text-xs font-extrabold text-teal-800 bg-teal-100 hover:bg-teal-200 px-2.5 sm:px-3 py-1.5 rounded-xl border border-teal-300 transition-all flex items-center gap-1 shadow-xs active:scale-95"
                                title="Bulk import students from Excel or text format Name (Class)"
                              >
                                <FileSpreadsheet className="w-3.5 h-3.5" />
                                <span>Excel / Bulk Import</span>
                              </button>

                              {/* 2. REMOVE STUDENT BUTTON */}
                              <button
                                onClick={() => {
                                  setRemoveSelectedRoomStudentIds(performanceSelectedStudents);
                                  setShowRemoveRoomStudentModal(true);
                                }}
                                className="text-xs font-extrabold text-rose-700 bg-rose-100 hover:bg-rose-200 px-2.5 sm:px-3 py-1.5 rounded-xl border border-rose-200 transition-all flex items-center gap-1 shadow-xs active:scale-95"
                                title="Remove student from this room"
                              >
                                <Minus className="w-3.5 h-3.5" />
                                <span>Remove</span>
                              </button>

                              {/* SELECT ALL / DESELECT ALL */}
                              <button
                                onClick={() => {
                                  if (selectedCount === allocatedRoomStudents.length) {
                                    setPerformanceSelectedStudents([]);
                                  } else {
                                    setPerformanceSelectedStudents(allocatedRoomStudents.map(s => s.id));
                                  }
                                }}
                                className="text-xs font-bold text-emerald-700 hover:text-emerald-900 bg-emerald-100/80 hover:bg-emerald-200 px-2.5 sm:px-3 py-1.5 rounded-xl border border-emerald-200 transition-colors"
                              >
                                {selectedCount === allocatedRoomStudents.length ? 'Deselect All' : 'Select All'}
                              </button>

                              {/* CHANGE */}
                              <button
                                onClick={() => {
                                  setSelectedRoom(null);
                                  setPerformanceSelectedStudents([]);
                                }}
                                className="text-xs font-bold text-slate-600 hover:text-slate-800 bg-white hover:bg-slate-100 px-2.5 sm:px-3 py-1.5 rounded-xl border border-slate-200 transition-colors shadow-2xs"
                              >
                                Change
                              </button>
                            </div>
                          </div>

                          {/* MAIN AREA: Student Checklist takes maximum height, showing all allocated students */}
                          <div className="flex-1 overflow-y-auto p-3 sm:p-4 flex flex-col gap-2.5 pb-36 touch-pan-y custom-scrollbar">
                            {allocatedRoomStudents.length > 4 && (
                              <p className="text-[10px] font-bold text-slate-400 text-center -mt-1 mb-0.5">
                                ↕️ Scroll up/down to see all {allocatedRoomStudents.length} students
                              </p>
                            )}
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
                                  className={`p-3.5 border rounded-2xl flex items-center justify-between transition-all active:scale-[0.99] text-left shadow-2xs ${isChecked
                                      ? 'bg-emerald-50/80 border-emerald-500 text-emerald-950 shadow-xs'
                                      : 'bg-white border-slate-200 text-slate-700 hover:border-slate-300 hover:bg-slate-50/50'
                                    }`}
                                >
                                  <div className="flex items-center gap-3">
                                    <div className={`w-8 h-8 rounded-xl flex items-center justify-center font-bold text-xs ${isChecked ? 'bg-emerald-600 text-white' : 'bg-slate-100 text-slate-600'
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
                                  <div className={`w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-colors ${isChecked ? 'bg-emerald-600 border-emerald-600' : 'border-slate-300 bg-white'
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
                                setRoomTallyCount(3);
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
                          className={`p-3 border rounded-xl flex items-center justify-between transition-colors ${performanceSelectedStudents.includes(student.id)
                              ? 'bg-[#1A365D]/10 border-[#1A365D] text-[#1A365D]'
                              : 'bg-white border-slate-200 text-slate-700 hover:border-[#1A365D]/50'
                            }`}
                        >
                          <span className="font-bold text-sm">{student.name}</span>
                          <div className={`w-5 h-5 rounded-md border-2 flex items-center justify-center ${performanceSelectedStudents.includes(student.id) ? 'bg-[#1A365D] border-[#1A365D]' : 'border-slate-300'
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
        ) : activeTab === 'phone_pass' ? (
          /* PHONE PASS DASHBOARD VIEW */
          <div className="flex-1 flex flex-col overflow-hidden bg-slate-50 relative font-sans">
            {/* Header Bar */}
            <div className="p-4 bg-white border-b border-slate-200 flex items-center justify-between shrink-0 shadow-xs flex-wrap gap-2">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-gradient-to-tr from-amber-500 to-emerald-600 text-white flex items-center justify-center font-bold shadow-md">
                  <Smartphone className="w-5 h-5" />
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <h2 className="text-[#1A365D] font-extrabold text-base uppercase leading-tight">School Phone Pass System</h2>
                    <span className="bg-emerald-100 text-emerald-800 text-[10px] font-black px-2 py-0.5 rounded-full uppercase">Digital Pass</span>
                  </div>
                  <p className="text-slate-500 text-xs font-medium">Security Office Mobile Issue & Return Tracker</p>
                </div>
              </div>

              {hasPermission('phone_pass_issue') && (
                <div className="flex items-center gap-2">
                  {hasPermission('admin_phone_pass') && phonePasses.length > 0 && (
                    <button
                      onClick={handleClearAllPhonePasses}
                      className="py-2.5 px-3.5 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl font-extrabold text-xs transition-all active:scale-95 flex items-center gap-1.5"
                      title="Delete all phone pass cards"
                    >
                      <Trash2 className="w-4 h-4" />
                      <span>Clear All Passes</span>
                    </button>
                  )}

                  <button
                    onClick={() => setShowMonthlyLeaveModal(true)}
                    className="py-2.5 px-3.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-extrabold text-xs shadow-md transition-all active:scale-95 flex items-center gap-1.5"
                  >
                    <Calendar className="w-4 h-4" />
                    <span>Monthly Leave</span>
                  </button>

                  <button
                    onClick={() => {
                      setPhonePassStep(1);
                      setPhonePassSelectedStudent(null);
                      setPhonePassReason('Calling Home');
                      setPhonePassAllowedMins(60);
                      setPhonePassCustomAllowedTime('');
                      setShowIssuePhonePassModal(true);
                    }}
                    className="py-2.5 px-4 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-extrabold text-xs shadow-md transition-all active:scale-95 flex items-center gap-1.5"
                  >
                    <Smartphone className="w-4 h-4" />
                    <Plus className="w-3.5 h-3.5 -ml-1" />
                    <span>Issue Phone Pass</span>
                  </button>
                </div>
              )}
            </div>

            {/* Search & Filter Controls */}
            <div className="p-3 sm:p-4 bg-white border-b border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3 shrink-0">
              <div className="relative w-full sm:w-80">
                <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={phonePassSearchQuery}
                  onChange={e => setPhonePassSearchQuery(e.target.value)}
                  placeholder="Search student name, class, reason..."
                  className="w-full py-2.5 pl-10 pr-4 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-500"
                />
              </div>

              {/* Status Filter Tabs */}
              <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 w-full sm:w-auto overflow-x-auto">
                {[
                  { key: 'all', label: `All (${phonePasses.length})` },
                  { key: 'ISSUED', label: `🔵 ISSUED (${phonePasses.filter(p => (p.status || 'ISSUED') === 'ISSUED').length})` },
                  { key: 'OUT', label: `🟢 OUT (${phonePasses.filter(p => p.status === 'OUT').length})` },
                  { key: 'IN', label: `✅ IN (${phonePasses.filter(p => p.status === 'IN').length})` },
                  { key: 'late', label: `⚠️ Late (${phonePasses.filter(p => p.isLate || (p.status === 'OUT' && new Date() > new Date(p.allowedUntil))).length})` }
                ].map(tab => (
                  <button
                    key={tab.key}
                    type="button"
                    onClick={() => setPhonePassFilterStatus(tab.key)}
                    className={`flex-1 sm:flex-none px-3 py-1.5 rounded-lg text-xs font-black uppercase transition-all whitespace-nowrap ${phonePassFilterStatus === tab.key
                        ? 'bg-white text-slate-800 shadow-xs'
                        : 'text-slate-500 hover:text-slate-800'
                      }`}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>
            </div>

            {/* Pass Cards Grid */}
            <div className="flex-1 overflow-y-auto p-3 sm:p-4 pb-28 custom-scrollbar">
              {(() => {
                const filtered = phonePasses.filter(p => {
                  const matchesSearch = !phonePassSearchQuery ||
                    p.studentName.toLowerCase().includes(phonePassSearchQuery.toLowerCase()) ||
                    p.studentClass.toLowerCase().includes(phonePassSearchQuery.toLowerCase()) ||
                    p.reason.toLowerCase().includes(phonePassSearchQuery.toLowerCase());

                  if (!matchesSearch) return false;

                  const isOverdue = p.status === 'OUT' && new Date() > new Date(p.allowedUntil);
                  const passStatus = p.status || 'ISSUED';

                  if (phonePassFilterStatus === 'ISSUED') return passStatus === 'ISSUED';
                  if (phonePassFilterStatus === 'OUT') return passStatus === 'OUT';
                  if (phonePassFilterStatus === 'IN') return passStatus === 'IN';
                  if (phonePassFilterStatus === 'late') return p.isLate || (passStatus === 'OUT' && isOverdue);
                  return true;
                });

                if (filtered.length === 0) {
                  return (
                    <div className="p-12 text-center bg-white rounded-3xl border border-dashed border-slate-200 flex flex-col items-center justify-center gap-3">
                      <div className="w-12 h-12 rounded-2xl bg-slate-100 text-slate-400 flex items-center justify-center">
                        <Smartphone className="w-6 h-6" />
                      </div>
                      <p className="text-slate-500 font-extrabold text-sm">No phone passes found.</p>
                      <p className="text-slate-400 text-xs">Issue a pass or try changing the search filter.</p>
                    </div>
                  );
                }

                return (
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                    {filtered.map(pass => {
                      const passStatus = pass.status || 'ISSUED';
                      const isOverdue = passStatus === 'OUT' && new Date() > new Date(pass.allowedUntil);
                      const lateMins = (new Date((pass.returnTime ? new Date(pass.returnTime) : new Date()) - new Date(pass.allowedUntil))).getTime() / 60000;
                      
                      const startDateObj = pass.startTime ? new Date(pass.startTime) : null;
                      const formattedStart = startDateObj ? startDateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true }) : '--:--';
                      const formattedStartDate = startDateObj ? startDateObj.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }) : '';
                      
                      const allowedDateObj = pass.allowedUntil ? new Date(pass.allowedUntil) : null;
                      const formattedAllowedTime = allowedDateObj ? allowedDateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true }) : '--:--';
                      const formattedAllowedDate = allowedDateObj ? allowedDateObj.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : '';

                      const returnDateObj = pass.returnTime ? new Date(pass.returnTime) : null;
                      const formattedReturnTime = returnDateObj ? returnDateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true }) : '';
                      const formattedReturnDate = returnDateObj ? returnDateObj.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }) : '';

                      return (
                        <div
                          key={pass.id}
                          className={`bg-white rounded-2xl border p-4 flex flex-col gap-3 shadow-xs transition-all relative overflow-hidden ${passStatus === 'ISSUED'
                              ? 'border-blue-200 ring-1 ring-blue-400/30'
                              : passStatus === 'OUT'
                                ? isOverdue
                                  ? 'border-rose-300 ring-2 ring-rose-500/20 bg-rose-50/20'
                                  : 'border-amber-200 ring-1 ring-amber-400/30'
                                : 'border-slate-200'
                            }`}
                        >
                          {/* Card Top: Student Name & Class */}
                          <div className="flex items-start justify-between gap-2 border-b border-slate-100 pb-2.5">
                            <div className="flex items-center gap-2.5">
                              <div className={`w-10 h-10 rounded-xl flex items-center justify-center font-black text-sm text-white ${passStatus === 'ISSUED'
                                  ? 'bg-blue-600'
                                  : passStatus === 'OUT'
                                    ? (isOverdue ? 'bg-rose-600' : 'bg-amber-500')
                                    : 'bg-emerald-600'
                                }`}>
                                {pass.studentName.charAt(0)}
                              </div>
                              <div>
                                <h4 className="font-black text-sm text-slate-800 leading-tight">{pass.studentName}</h4>
                                <div className="flex items-center gap-1.5 mt-0.5">
                                  <span className="font-extrabold text-[10px] bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md uppercase">
                                    Class {pass.studentClass}
                                  </span>
                                  {(() => {
                                    const matchingStudent = students.find(s => String(s.id) === String(pass.studentId));
                                    return matchingStudent?.registerNumber ? (
                                      <span className="font-mono font-black text-[10px] bg-indigo-50 text-indigo-700 px-1.5 py-0.5 rounded-md border border-indigo-100">
                                        #{matchingStudent.registerNumber}
                                      </span>
                                    ) : null;
                                  })()}
                                </div>
                              </div>
                            </div>

                            {/* Status Badge */}
                            {passStatus === 'ISSUED' ? (
                              <span className="px-2.5 py-1 rounded-full text-[10px] font-black uppercase tracking-wider bg-blue-100 text-blue-800 flex items-center gap-1">
                                <span className="w-2 h-2 rounded-full bg-blue-600 animate-ping" />
                                ISSUED (Ready)
                              </span>
                            ) : passStatus === 'OUT' ? (
                              <span className={`px-2.5 py-1 rounded-full text-[10px] font-black uppercase tracking-wider flex items-center gap-1 ${isOverdue ? 'bg-rose-100 text-rose-700 animate-pulse' : 'bg-amber-100 text-amber-800'
                                }`}>
                                <span className={`w-2 h-2 rounded-full ${isOverdue ? 'bg-rose-600' : 'bg-amber-500 animate-ping'}`} />
                                {isOverdue ? 'OVERDUE (OUT)' : 'OUT (In Use)'}
                              </span>
                            ) : (
                              <span className="px-2.5 py-1 rounded-full text-[10px] font-black uppercase tracking-wider bg-emerald-100 text-emerald-800 flex items-center gap-1">
                                <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                                IN (Returned)
                              </span>
                            )}
                          </div>

                          {/* Reason & Phone details */}
                          <div className="text-xs flex flex-col gap-1 bg-slate-50 p-2.5 rounded-xl border border-slate-100 font-semibold text-slate-600">
                            <div><span className="font-bold text-slate-400">Reason:</span> <span className="font-extrabold text-slate-800">{pass.reason}</span></div>
                            {pass.phoneModel && (
                              <div><span className="font-bold text-slate-400">Phone:</span> <span className="font-extrabold text-slate-700">{pass.phoneModel}</span></div>
                            )}
                          </div>

                          {/* Time details */}
                          <div className="grid grid-cols-2 gap-2 text-[11px] font-bold">
                            <div className="bg-slate-100 p-2 rounded-xl">
                              <span className="text-[9px] font-extrabold text-slate-400 uppercase block">
                                {passStatus === 'ISSUED' ? 'Issued Time' : 'Pickup Time (OUT)'}
                              </span>
                              <span className="text-slate-800 font-extrabold block">{formattedStart}</span>
                              {formattedStartDate && (
                                <span className="text-[10px] font-bold text-slate-500 block">{formattedStartDate}</span>
                              )}
                            </div>
                            <div className="bg-slate-100 p-2 rounded-xl">
                              <span className="text-[9px] font-extrabold text-slate-400 uppercase block">Return Date & Time</span>
                              <span className="text-slate-800 font-extrabold block">{formattedAllowedTime}</span>
                              {formattedAllowedDate && (
                                <span className="text-[10px] font-black text-emerald-700 block">{formattedAllowedDate}</span>
                              )}
                            </div>
                          </div>

                          {/* Return details & Submission Time */}
                          {passStatus === 'IN' && (
                            <div className="flex items-center justify-between text-xs bg-emerald-50 border border-emerald-100 p-2.5 rounded-xl text-emerald-900 font-bold">
                              <span>Submitted at:</span>
                              <span className="font-black">{formattedReturnTime} ({formattedReturnDate})</span>
                            </div>
                          )}

                          {/* Late warning badge */}
                          {(pass.isLate || isOverdue) && (
                            <div className="bg-rose-100 border border-rose-200 text-rose-800 p-2 rounded-xl text-xs font-black flex items-center justify-between">
                              <span className="flex items-center gap-1">
                                <AlertTriangle className="w-3.5 h-3.5 text-rose-600" />
                                {passStatus === 'IN' ? 'Returned LATE' : 'OVERDUE'}
                              </span>
                              <span>{lateMins > 0 ? `+${Math.ceil(lateMins)} Mins` : 'LATE'}</span>
                            </div>
                          )}

                          {/* Interactive OUT / IN Toggle Button & Admin Delete Action */}
                          <div className="pt-1 flex items-center gap-2">
                            <div className="flex-1">
                              {passStatus === 'ISSUED' ? (
                                <button
                                  onClick={() => handlePickupPhonePass(pass.id)}
                                  className="w-full py-2.5 px-4 bg-amber-500 hover:bg-amber-600 text-white rounded-xl font-extrabold text-xs shadow-md transition-all active:scale-95 flex items-center justify-center gap-2"
                                >
                                  <Smartphone className="w-4 h-4" />
                                  <span>Click OUT (Take Phone from Security)</span>
                                </button>
                              ) : passStatus === 'OUT' ? (
                                <button
                                  onClick={() => handleReturnPhonePass(pass.id)}
                                  className="w-full py-2.5 px-4 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-extrabold text-xs shadow-md transition-all active:scale-95 flex items-center justify-center gap-2"
                                >
                                  <Check className="w-4 h-4" />
                                  <span>Click IN (Return Phone to Security)</span>
                                </button>
                              ) : (
                                <div className="w-full py-2 px-4 bg-slate-100 text-slate-500 rounded-xl font-extrabold text-xs text-center border border-slate-200">
                                  Pass Completed (IN)
                                </div>
                              )}
                            </div>

                            {hasPermission('admin_phone_pass') && (
                              <button
                                onClick={() => handleDeletePhonePass(pass.id, pass.studentName)}
                                className="p-2.5 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-xl font-bold transition-all active:scale-95 flex items-center justify-center shrink-0"
                                title="Delete Phone Pass (Admin Only)"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            )}
                          </div>

                        </div>
                      );
                    })}
                  </div>
                );
              })()}
            </div>
          </div>
        ) : (
          /* SPREADSHEET VIEW (RBAC Protected) */
          <div className="flex-1 flex flex-col overflow-hidden bg-[#F8F9FA]">
            {/* Top Sub-tabs & Settings Bar (Admin Only) */}
            {isAdminAuthenticated ? (
              <div className="flex items-center gap-1 p-2 bg-white border-b border-slate-200 shrink-0">
                {hasPermission('admin_sheet') && (
                  <button
                    onClick={() => setAdminSubTab('sheet')}
                    className={`flex-1 py-2 rounded-xl text-xs font-extrabold uppercase tracking-wider transition-all ${adminSubTab === 'sheet' ? 'bg-[#1A365D] text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100'}`}
                  >Score Sheet</button>
                )}

                {hasPermission('admin_phone_pass') && (
                  <button
                    onClick={() => setAdminSubTab('phone_pass')}
                    className={`flex-1 py-2 rounded-xl text-xs font-extrabold uppercase tracking-wider transition-all flex items-center justify-center gap-1 ${adminSubTab === 'phone_pass' ? 'bg-emerald-700 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100'}`}
                  >
                    <Smartphone className="w-3.5 h-3.5" />
                    <span>Phone Pass</span>
                  </button>
                )}

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
            ) : hasPermission('user_sheet') ? (
              <div className="flex items-center justify-between p-2 bg-white border-b border-slate-200 shrink-0">
                <span className="text-xs font-extrabold uppercase tracking-wider text-[#1A365D] pl-1">📋 Score Sheet</span>
                <button
                  onClick={handleLogout}
                  className="p-2 rounded-xl hover:bg-red-50 text-slate-400 hover:text-red-500 transition-all"
                  title="Logout"
                ><Lock className="w-4 h-4" /></button>
              </div>
            ) : null}

            {adminSubTab === 'phone_pass' && hasPermission('admin_phone_pass') ? (
              /* ─── PHONE PASS MANAGEMENT UI IN ADMIN PANEL ─── */
              <div className="flex-1 flex flex-col overflow-hidden bg-slate-50 font-sans">
                {/* Top Bar with Title, School vs Home Filter & Issue Pass Button */}
                <div className="p-3.5 bg-white border-b border-slate-200 flex flex-col gap-3 shrink-0">
                  {/* Header Title & Issue Action Row */}
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2.5">
                      <div className="w-9 h-9 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-bold shadow-xs shrink-0">
                        <Smartphone className="w-5 h-5" />
                      </div>
                      <div>
                        <h3 className="text-sm font-black text-slate-800 uppercase leading-tight">Phone Pass Management</h3>
                        <p className="text-[11px] font-semibold text-slate-500">Student phone registry & issuance</p>
                      </div>
                    </div>

                    {hasPermission('phone_pass_issue') && (
                      <div className="flex items-center gap-2">
                        {hasPermission('admin_phone_pass') && phonePasses.length > 0 && (
                          <button
                            onClick={handleClearAllPhonePasses}
                            className="py-2 px-3 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl font-extrabold text-xs transition-all active:scale-95 flex items-center gap-1 shrink-0 whitespace-nowrap"
                            title="Delete all phone pass cards"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            <span>Clear All</span>
                          </button>
                        )}

                        <button
                          onClick={() => setShowMonthlyLeaveModal(true)}
                          className="py-2 px-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-extrabold text-xs shadow-sm transition-all active:scale-95 flex items-center gap-1 shrink-0 whitespace-nowrap"
                        >
                          <Calendar className="w-3.5 h-3.5" />
                          <span>Monthly Leave</span>
                        </button>

                        <button
                          onClick={() => {
                            setPhonePassStep(1);
                            setPhonePassSelectedStudent(null);
                            setShowIssuePhonePassModal(true);
                          }}
                          className="py-2 px-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-extrabold text-xs shadow-sm transition-all active:scale-95 flex items-center gap-1 shrink-0 whitespace-nowrap"
                        >
                          <Smartphone className="w-3.5 h-3.5" />
                          <Plus className="w-3.5 h-3.5 -ml-0.5" />
                          <span>+ Issue Pass</span>
                        </button>
                      </div>
                    )}
                  </div>

                  {/* School vs Home Toggle Filter Segment */}
                  <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200 w-full">
                    <button
                      type="button"
                      onClick={() => setPhonePassStudentTypeFilter('school')}
                      className={`flex-1 py-1.5 px-3 rounded-lg text-xs font-black uppercase transition-all flex items-center justify-center gap-1.5 ${phonePassStudentTypeFilter === 'school' ? 'bg-emerald-600 text-white shadow-xs' : 'text-slate-600 hover:text-slate-800'
                        }`}
                    >
                      <School className="w-3.5 h-3.5" />
                      <span>School ({students.filter(s => (s.phoneType || 'school') === 'school').length})</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setPhonePassStudentTypeFilter('home')}
                      className={`flex-1 py-1.5 px-3 rounded-lg text-xs font-black uppercase transition-all flex items-center justify-center gap-1.5 ${phonePassStudentTypeFilter === 'home' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-600 hover:text-slate-800'
                        }`}
                    >
                      <UserCheck className="w-3.5 h-3.5" />
                      <span>Home ({students.filter(s => s.phoneType === 'home').length})</span>
                    </button>
                  </div>
                </div>

                {/* Search and Class Filter Bar */}
                <div className="p-3 bg-white border-b border-slate-200 flex flex-wrap items-center justify-between gap-2 shrink-0">
                  <div className="relative flex-1 min-w-[200px]">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={phonePassStudentSearch}
                      onChange={e => setPhonePassStudentSearch(e.target.value)}
                      placeholder="Search student by name..."
                      className="w-full py-2 pl-9 pr-3 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  <select
                    value={phonePassClassFilter}
                    onChange={e => setPhonePassClassFilter(e.target.value)}
                    className="py-2 px-3 bg-slate-50 border border-slate-200 rounded-xl font-extrabold text-xs text-slate-800 uppercase focus:outline-none focus:border-emerald-500"
                  >
                    <option value="all">All Classes</option>
                    {CLASSES.map(c => (
                      <option key={c} value={c}>Class {c.toUpperCase()}</option>
                    ))}
                  </select>
                </div>

                {/* Student Registry Table */}
                <div className="flex-1 overflow-y-auto p-3 sm:p-4 custom-scrollbar">
                  <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-100 text-slate-500 font-extrabold uppercase text-[10px] tracking-wider border-b border-slate-200">
                        <tr>
                          <th className="p-3">#</th>
                          <th className="p-3">Reg No</th>
                          <th className="p-3">Student Name</th>
                          <th className="p-3">Class</th>
                          <th className="p-3">Category</th>
                          <th className="p-3">Phone Model</th>
                          <th className="p-3">Pass Status</th>
                          <th className="p-3 text-right">Action</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-semibold text-slate-700">
                        {students
                          .filter(s => {
                            const matchType = (s.phoneType || 'school') === phonePassStudentTypeFilter;
                            const matchSearch = !phonePassStudentSearch ||
                              s.name.toLowerCase().includes(phonePassStudentSearch.toLowerCase()) ||
                              (s.registerNumber && s.registerNumber.includes(phonePassStudentSearch));
                            const matchClass = phonePassClassFilter === 'all' || s.class.toLowerCase() === phonePassClassFilter.toLowerCase();
                            return matchType && matchSearch && matchClass;
                          })
                          .map((st, idx) => {
                            const activePass = phonePasses.find(p => String(p.studentId) === String(st.id) && p.status === 'OUT');

                            return (
                              <tr key={st.id} className="hover:bg-slate-50 transition-colors">
                                <td className="p-3 font-mono text-[10px] text-slate-400">#{idx + 1}</td>
                                <td className="p-3">
                                  <input
                                    type="text"
                                    maxLength={5}
                                    defaultValue={st.registerNumber || ''}
                                    onBlur={e => {
                                      const val = e.target.value.trim();
                                      if (val !== (st.registerNumber || '')) {
                                        handleUpdateStudentPhoneDetails(st.id, st.phoneType || 'school', st.phoneModel || '', val);
                                      }
                                    }}
                                    placeholder="5-digit"
                                    className="py-1 px-2 bg-slate-50 border border-slate-200 rounded-lg text-xs font-mono font-extrabold text-slate-800 focus:outline-none focus:border-emerald-500 w-20 tracking-wider"
                                  />
                                </td>
                                <td className="p-3 font-extrabold text-slate-800">{st.name}</td>
                                <td className="p-3">
                                  <span className="font-extrabold text-[10px] bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md uppercase">
                                    {st.class}
                                  </span>
                                </td>
                                <td className="p-3">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const newType = (st.phoneType || 'school') === 'school' ? 'home' : 'school';
                                      handleUpdateStudentPhoneDetails(st.id, newType, st.phoneModel || '', st.registerNumber || '');
                                    }}
                                    className={`px-2.5 py-1 rounded-full text-[10px] font-black uppercase transition-all flex items-center gap-1 w-fit cursor-pointer ${(st.phoneType || 'school') === 'school'
                                        ? 'bg-emerald-100 text-emerald-800 hover:bg-emerald-200'
                                        : 'bg-blue-100 text-blue-800 hover:bg-blue-200'
                                      }`}
                                  >
                                    {(st.phoneType || 'school') === 'school' ? '🏫 School' : '🏠 Home'}
                                  </button>
                                </td>
                                <td className="p-3">
                                  <input
                                    type="text"
                                    defaultValue={st.phoneModel || ''}
                                    onBlur={e => {
                                      if (e.target.value !== (st.phoneModel || '')) {
                                        handleUpdateStudentPhoneDetails(st.id, st.phoneType || 'school', e.target.value, st.registerNumber || '');
                                      }
                                    }}
                                    placeholder="Enter phone details..."
                                    className="py-1 px-2.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-500 w-44"
                                  />
                                </td>
                                <td className="p-3">
                                  {activePass ? (
                                    <span className="font-black text-[10px] bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full uppercase flex items-center gap-1 w-fit animate-pulse">
                                      <span className="w-1.5 h-1.5 rounded-full bg-amber-600" />
                                      OUT (Active)
                                    </span>
                                  ) : (
                                    <span className="font-bold text-[10px] text-slate-400 uppercase">Clear</span>
                                  )}
                                </td>
                                <td className="p-3 text-right">
                                  <button
                                    disabled={(st.phoneType || 'school') === 'home' || !st.phoneModel || !st.phoneModel.trim() || !!activePass || !hasPermission('phone_pass_issue')}
                                    onClick={() => {
                                      setPhonePassSelectedStudent(st);
                                      setPhonePassStep(2);
                                      setShowIssuePhonePassModal(true);
                                    }}
                                    title={!st.phoneModel || !st.phoneModel.trim() ? "Add Phone Name first to issue pass" : ""}
                                    className="py-1.5 px-3 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 disabled:cursor-not-allowed text-white rounded-xl font-extrabold text-[11px] transition-all"
                                  >
                                    + Issue Pass
                                  </button>
                                </td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            ) : isSuperAdmin && adminSubTab === 'users' ? (
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
                          <span className={`text-[10px] font-extrabold px-2 py-0.5 rounded-full uppercase ${u.role === 'super_admin' ? 'bg-amber-100 text-amber-800 border border-amber-300' :
                              u.role === 'admin' ? 'bg-purple-100 text-purple-700' :
                                (Array.isArray(u.permissions) && (u.permissions.includes('admin_sheet') || u.permissions.includes('admin_phone_pass'))) ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-slate-100 text-slate-600'
                            }`}>
                            {u.role === 'super_admin' ? 'Super Admin' :
                              u.role === 'admin' ? 'Admin' :
                                (Array.isArray(u.permissions) && (u.permissions.includes('admin_sheet') || u.permissions.includes('admin_phone_pass'))) ? 'Partial Admin' : 'User'}
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
                <div className={`grid ${isAdminAuthenticated ? 'grid-cols-5' : 'grid-cols-2'} gap-2`}>
                  <button
                    onClick={() => {
                      setDownloadSelectedClasses(CLASSES);
                      setDateModalNextAction('ADMIN_REPORT');
                      setShowIRDateModal(true);
                    }}
                    className="flex items-center justify-center gap-1.5 py-2.5 px-1 rounded-xl font-extrabold text-[11px] bg-emerald-700 hover:bg-emerald-800 text-white shadow-xs active:scale-[0.98] transition-all"
                  >
                    <Download className="w-3.5 h-3.5 shrink-0" />
                    Download
                  </button>

                  <button
                    onClick={() => setShowActivitiesReportModal(true)}
                    className="flex items-center justify-center gap-1.5 py-2.5 px-1 rounded-xl font-extrabold text-[11px] bg-teal-600 hover:bg-teal-700 text-white shadow-xs active:scale-[0.98] transition-all"
                  >
                    <FileText className="w-3.5 h-3.5 shrink-0" />
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
                    <thead className="sticky top-0 z-20 bg-slate-100 shadow-xs">
                      {/* Outer Class Header Row */}
                      <tr className="border-b border-slate-200 bg-slate-100 text-center font-bold">
                        <th className="border-r border-slate-200 p-2 text-slate-400 font-mono text-[10px] w-8 bg-slate-100 sticky top-0 z-20">
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
                        <th className="border-r border-slate-200 p-2 text-slate-400 font-mono text-[10px] w-8 bg-slate-100 sticky top-0 z-20">#</th>
                        <th colSpan="13" className="p-2 text-[#1A365D] font-extrabold uppercase tracking-widest text-xs bg-[#1A365D]/10 sticky top-0 z-20">
                          Class {adminClass.toUpperCase()}
                        </th>
                      </tr>

                      {/* Table headers */}
                      <tr className="border-b border-slate-200 bg-slate-100 text-slate-600 uppercase text-[10px] font-extrabold">
                        <th className="border-r border-slate-200 p-2 text-center bg-slate-100 sticky top-[33px] z-20"></th>
                        <th className="border-r border-slate-200 p-2 text-center font-mono text-[10px] bg-slate-100 sticky top-[33px] z-20">ROW</th>
                        <th className="border-r border-slate-200 p-2.5 font-bold min-w-[260px] w-[260px] text-left bg-slate-100 sticky top-[33px] z-20">STUDENT NAME</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-amber-700 bg-slate-100 sticky top-[33px] z-20 min-w-[75px]">STARS</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-sky-700 bg-slate-100 sticky top-[33px] z-20 min-w-[75px]">TALLIES</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-[#1A365D] bg-slate-100 sticky top-[33px] z-20 min-w-[75px]">TOTAL</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-purple-700 bg-slate-100 sticky top-[33px] z-20 min-w-[75px]">GRADE</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-orange-600 bg-slate-100 sticky top-[33px] z-20 min-w-[95px]">N&O TALLY</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-[#1A365D] bg-slate-100 sticky top-[33px] z-20 min-w-[75px]">TOTAL</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-orange-700 bg-slate-100 sticky top-[33px] z-20 min-w-[95px]">N&O GRADE</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-sky-600 bg-slate-100 sticky top-[33px] z-20 min-w-[105px]">DIARY TALLIES</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-cyan-600 bg-slate-100 sticky top-[33px] z-20 min-w-[75px]">SHEETS</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-rose-600 bg-slate-100 sticky top-[33px] z-20 min-w-[75px]">FINE</th>
                        <th className="border-r border-slate-200 p-2.5 text-center font-bold text-[#1A365D] bg-slate-100 sticky top-[33px] z-20 min-w-[75px]">TOTAL</th>
                        <th className="p-2.5 text-center font-bold text-indigo-700 bg-slate-100 sticky top-[33px] z-20 min-w-[120px]">ATTITUDE GRADE</th>
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
                              <td className="border-r border-slate-200 p-1 font-semibold text-slate-800 min-w-[260px] w-[260px]">
                                <input
                                  type="text"
                                  value={student.name}
                                  readOnly={!isAdminAuthenticated}
                                  onChange={(e) => updateStudentField(student.id, 'name', e.target.value)}
                                  className={`bg-transparent text-left w-full focus:outline-none py-1 px-2 rounded font-semibold text-xs text-slate-800 min-w-[240px] whitespace-nowrap overflow-visible ${isAdminAuthenticated ? 'focus:bg-slate-100 cursor-pointer' : 'cursor-default'}`}
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
                                  className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-extrabold ${finalTotal > 0 ? 'text-emerald-700' : finalTotal < 0 ? 'text-rose-600' : 'text-slate-500'
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
                                  className={`bg-transparent text-center w-full focus:outline-none py-1 px-1 rounded font-extrabold uppercase ${finalTotal >= 20 ? 'text-emerald-700' : finalTotal >= 7 ? 'text-emerald-500' : finalTotal >= 0 ? 'text-amber-500' : finalTotal >= -6 ? 'text-orange-500' : finalTotal >= -20 ? 'text-rose-500' : 'text-rose-700'
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
                                  value={getFineCount(student)}
                                  onChange={(e) => {
                                    const val = Math.max(0, parseInt(e.target.value) || 0);
                                    updateStudentField(student.id, 'fine', val);
                                    updateStudentField(student.id, 'fineCount', val);
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

        {/* Persistent Bottom Tab Navigation Bar */}
        <footer className="h-16 border-t border-slate-200 bg-white flex items-center justify-around px-2 shadow-[0_-4px_16px_rgba(0,0,0,0.03)] shrink-0 z-10 font-sans">
          {hasPermission('scoring') && (
            <button
              onClick={() => {
                setActiveTab('scoring');
                setSelectedClass(null);
                setSelectedStudentIds([]);
                setIsScoring(false);
                setSearchQuery('');
              }}
              className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 ${activeTab === 'scoring'
                  ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]'
                  : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
                }`}
            >
              <Sparkles className={`w-4 h-4 ${activeTab === 'scoring' ? 'text-amber-300' : ''}`} />
              <span className="text-[10px] tracking-wide">Scoring</span>
            </button>
          )}

          {hasPermission('mentor') && (
            <button
              onClick={() => {
                setActiveTab('mentor');
                setSaveStatus('');
              }}
              className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 ${activeTab === 'mentor'
                  ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]'
                  : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
                }`}
            >
              <UserCheck className={`w-4 h-4 ${activeTab === 'mentor' ? 'text-sky-300' : ''}`} />
              <span className="text-[10px] tracking-wide">Mentor</span>
            </button>
          )}

          {hasPermission('performance') && (
            <button
              onClick={() => {
                setActiveTab('performance');
                setSaveStatus('');
              }}
              className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 ${activeTab === 'performance'
                  ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]'
                  : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
                }`}
            >
              <TrendingUp className={`w-4 h-4 ${activeTab === 'performance' ? 'text-emerald-300' : ''}`} />
              <span className="text-[10px] tracking-wide">Performance</span>
            </button>
          )}

          {(hasPermission('phone_pass') || hasPermission('phone_pass_issue') || hasPermission('admin_phone_pass')) && (
            <button
              onClick={() => {
                setActiveTab('phone_pass');
                setSaveStatus('');
              }}
              className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 relative ${activeTab === 'phone_pass'
                  ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]'
                  : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
                }`}
            >
              <div className="relative">
                <Smartphone className={`w-4 h-4 ${activeTab === 'phone_pass' ? 'text-amber-300' : ''}`} />
                {phonePasses.filter(p => p.status === 'OUT').length > 0 && (
                  <span className="absolute -top-1.5 -right-2.5 bg-rose-500 text-white text-[9px] font-black w-4 h-4 rounded-full flex items-center justify-center shadow-xs">
                    {phonePasses.filter(p => p.status === 'OUT').length}
                  </span>
                )}
              </div>
              <span className="text-[10px] tracking-wide">Phone Pass</span>
            </button>
          )}

          {(hasPermission('user_sheet') || hasPermission('admin_sheet') || isAdminAuthenticated) && (
            <button
              onClick={() => {
                setActiveTab('admin');
                setSaveStatus('');
              }}
              className={`flex-1 py-1.5 px-1 rounded-xl flex flex-col items-center justify-center gap-0.5 transition-all duration-200 ${activeTab === 'admin'
                  ? 'bg-[#1A365D] text-white shadow-xs font-extrabold scale-[1.02]'
                  : 'text-slate-500 hover:text-[#1A365D] hover:bg-slate-100 font-semibold'
                }`}
            >
              <Table className={`w-4 h-4 ${activeTab === 'admin' ? 'text-purple-300' : ''}`} />
              <span className="text-[10px] tracking-wide">
                {isAdminAuthenticated ? 'Admin Sheet' : 'Score Sheet'}
              </span>
            </button>
          )}
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
                  className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${addStudentMethod === 'single'
                      ? 'bg-[#1A365D] text-white shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                    }`}
                >
                  Add Single Student
                </button>
                <button
                  type="button"
                  onClick={() => setAddStudentMethod('excel')}
                  className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${addStudentMethod === 'excel'
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
                        <span className={`font-mono font-extrabold px-2 py-0.5 rounded-full text-xs ${entry.type === 'star'
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
                        className={`flex items-center gap-2 p-2.5 rounded-lg border cursor-pointer transition-all ${isChecked
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
                          className={`w-4 h-4 rounded border flex items-center justify-center transition-all shrink-0 ${isChecked
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
                        className={`flex items-center gap-2 p-2.5 rounded-lg border cursor-pointer transition-all ${isChecked
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
                          className={`w-4 h-4 rounded border flex items-center justify-center transition-all shrink-0 ${isChecked
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

        {/* Add Class Mentor Modal */}
        {showAddClassMentorModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
            <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl border border-slate-200 overflow-hidden" onClick={e => e.stopPropagation()}>
              <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
                <div className="flex items-center gap-2 text-[#1A365D]">
                  <School className="w-5 h-5" />
                  <h2 className="text-sm font-extrabold tracking-wider">Add Class Mentor</h2>
                </div>
                <button
                  onClick={() => setShowAddClassMentorModal(false)}
                  className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
              <form onSubmit={handleAddClassMentor} className="p-5 flex flex-col gap-4">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                    1. Mentor Name
                  </label>
                  <input
                    type="text"
                    required
                    value={newClassMentor.name}
                    onChange={(e) => setNewClassMentor({ ...newClassMentor, name: e.target.value })}
                    className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:border-[#1A365D] focus:bg-white transition-colors"
                    placeholder="e.g. dddd"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                    2. Select Class (Show All Classes)
                  </label>
                  <select
                    required
                    value={newClassMentor.classAssigned}
                    onChange={(e) => setNewClassMentor({ ...newClassMentor, classAssigned: e.target.value })}
                    className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:border-[#1A365D] focus:bg-white transition-colors uppercase"
                  >
                    <option value="">-- Select Class --</option>
                    {CLASSES.map(c => (
                      <option key={c} value={c}>Class {c.toUpperCase()}</option>
                    ))}
                  </select>
                </div>

                <div className="pt-2">
                  <button
                    type="submit"
                    disabled={!newClassMentor.name.trim() || !newClassMentor.classAssigned.trim()}
                    className="w-full py-3 bg-[#1A365D] hover:bg-[#2A4365] text-white rounded-xl font-extrabold text-xs shadow-md transition-colors active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    Create Class Mentor
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* Delete Class Mentor Modal */}
        {showDeleteClassMentorModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[100] animate-fade-in">
            <div className="bg-white rounded-2xl w-full max-w-sm shadow-2xl border border-slate-200 overflow-hidden" onClick={e => e.stopPropagation()}>
              <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
                <div className="flex items-center gap-2 text-rose-600">
                  <Trash2 className="w-5 h-5" />
                  <h2 className="text-sm font-extrabold tracking-wider">Remove Class Mentor</h2>
                </div>
                <button
                  onClick={() => setShowDeleteClassMentorModal(false)}
                  className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
              <form onSubmit={handleDeleteClassMentor} className="p-5 flex flex-col gap-4">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide block mb-1">
                    Select Class Mentor to Remove
                  </label>
                  <div className="flex flex-col gap-2 max-h-[200px] overflow-y-auto pr-1">
                    {classMentors.length > 0 ? classMentors.map(cm => (
                      <button
                        key={cm.id}
                        type="button"
                        onClick={() => setSelectedClassMentorForDelete(cm.id)}
                        className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${selectedClassMentorForDelete === cm.id
                            ? 'border-rose-500 bg-rose-500/5 ring-1 ring-rose-500/20 shadow-sm'
                            : 'border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50'
                          }`}
                      >
                        <div>
                          <p className={`text-xs font-bold ${selectedClassMentorForDelete === cm.id ? 'text-rose-700' : 'text-slate-700'}`}>
                            {cm.name}
                          </p>
                          <p className="text-[10px] font-bold text-indigo-600 uppercase">
                            Class {cm.classAssigned.toUpperCase()}
                          </p>
                        </div>
                        {selectedClassMentorForDelete === cm.id && (
                          <CheckCircle2 className="w-4 h-4 text-rose-500" />
                        )}
                      </button>
                    )) : (
                      <p className="text-center text-slate-500 text-xs py-4">No Class Mentors to remove.</p>
                    )}
                  </div>
                </div>

                <div className="pt-2">
                  <button
                    type="submit"
                    disabled={!selectedClassMentorForDelete}
                    className="w-full py-3 bg-rose-600 hover:bg-rose-700 text-white rounded-xl font-extrabold text-xs shadow-md transition-colors active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    Remove Selected Class Mentor
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
                    onChange={(e) => setNewMentor({ ...newMentor, name: e.target.value })}
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
                    onChange={(e) => setNewMentor({ ...newMentor, roomNumber: e.target.value })}
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
                        className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${selectedMentorForDelete === m.id
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
                    onChange={(e) => setPerformanceSubmitData({ ...performanceSubmitData, count: e.target.value })}
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
                      setPerformanceSubmitData({ ...performanceSubmitData, reason: e.target.value });
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
                                setPerformanceSubmitData({ ...performanceSubmitData, reason: item.reason, count: item.count });
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
                    className={`w-full py-3 text-white rounded-xl font-extrabold text-xs shadow-md transition-colors active:scale-[0.98] ${performanceSubmitData.type === 'star' ? 'bg-amber-500 hover:bg-amber-600' : 'bg-rose-500 hover:bg-rose-600'
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
                        className={`flex items-center justify-between p-3 border rounded-xl transition-all text-left ${isSelected
                            ? 'border-rose-600 bg-rose-50 ring-1 ring-rose-600/20 shadow-sm'
                            : 'border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50'
                          }`}
                      >
                        <div className="flex items-center gap-3">
                          <div className={`w-9 h-9 rounded-full flex items-center justify-center text-[11px] font-extrabold transition-colors ${isSelected ? 'bg-rose-600 text-white' : 'bg-slate-100 text-[#1A365D]'
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

        {/* IR Date & Student Selection Modal */}
        {showIRDateModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[105] animate-fade-in" onClick={e => e.stopPropagation()}>
            <div className="bg-white rounded-2xl w-full max-w-4xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col scale-in max-h-[90vh]">
              <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50 shrink-0">
                <div className="flex items-center gap-2 text-[#1A365D]">
                  <FileText className="w-5 h-5 text-indigo-600" />
                  <div>
                    <h2 className="text-sm font-extrabold tracking-wider uppercase">Individual Report (IR) Filter</h2>
                    <p className="text-[11px] text-slate-500 font-medium">Configure date range and select student(s) to generate report</p>
                  </div>
                </div>
                <button onClick={() => setShowIRDateModal(false)} className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors">
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="p-6 flex-1 overflow-y-auto bg-slate-50/50 flex flex-col gap-6">
                {/* 3 Column Inputs: From Date, End Date, Student Multi-Select */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {/* Column 1: From Date */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-col gap-2">
                    <label className="text-xs font-extrabold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <Calendar className="w-4 h-4 text-indigo-600" />
                      From Date
                    </label>
                    <input
                      type="date"
                      value={irFromDate}
                      onChange={(e) => setIrFromDate(e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all font-semibold text-xs text-slate-800"
                    />
                    <p className="text-[10px] text-slate-400 font-medium">Start date of event range</p>
                  </div>

                  {/* Column 2: End Date */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-col gap-2">
                    <label className="text-xs font-extrabold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <Calendar className="w-4 h-4 text-indigo-600" />
                      End Date
                    </label>
                    <input
                      type="date"
                      value={irToDate}
                      onChange={(e) => setIrToDate(e.target.value)}
                      className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all font-semibold text-xs text-slate-800"
                    />
                    <p className="text-[10px] text-slate-400 font-medium">End date of event range</p>
                  </div>

                  {/* Column 3: Select Student or Multi Selection */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-col gap-2.5">
                    <div className="flex items-center justify-between">
                      <label className="text-xs font-extrabold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                        <UserCheck className="w-4 h-4 text-indigo-600" />
                        Select Student(s)
                      </label>
                      <span className="text-[10px] font-extrabold text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded-full border border-indigo-100">
                        {selectedIRStudentIds.length} / {irAssignedStudents.length}
                      </span>
                    </div>

                    {/* Multi-Select Controls: Select All / Deselect All */}
                    <div className="flex items-center justify-between gap-2 border-b border-slate-100 pb-2">
                      <button
                        type="button"
                        onClick={() => {
                          if (selectedIRStudentIds.length === irAssignedStudents.length) {
                            setSelectedIRStudentIds([]);
                          } else {
                            setSelectedIRStudentIds(irAssignedStudents.map(s => s.id));
                          }
                        }}
                        className="text-[11px] font-bold text-indigo-600 hover:text-indigo-800 hover:underline transition-all"
                      >
                        {selectedIRStudentIds.length === irAssignedStudents.length ? 'Deselect All' : 'Select All'}
                      </button>
                      <input
                        type="text"
                        placeholder="Search..."
                        value={irStudentSearch}
                        onChange={(e) => setIrStudentSearch(e.target.value)}
                        className="px-2.5 py-1 bg-slate-50 border border-slate-200 rounded-lg text-[11px] font-semibold text-slate-700 w-28 focus:outline-none focus:border-indigo-500"
                      />
                    </div>

                    {/* Student Checkboxes List */}
                    <div className="max-h-40 overflow-y-auto flex flex-col gap-1 pr-1 border border-slate-100 rounded-lg p-1.5 bg-slate-50/50">
                      {irAssignedStudents.filter(s => (s.name || '').toLowerCase().includes(irStudentSearch.toLowerCase())).length > 0 ? (
                        irAssignedStudents
                          .filter(s => (s.name || '').toLowerCase().includes(irStudentSearch.toLowerCase()))
                          .map(student => {
                            const isChecked = selectedIRStudentIds.includes(student.id);
                            return (
                              <label
                                key={student.id}
                                className={`flex items-center justify-between p-2 rounded-lg text-xs font-semibold cursor-pointer transition-colors ${isChecked ? 'bg-indigo-50 text-indigo-900 border border-indigo-200' : 'bg-white text-slate-700 hover:bg-slate-100 border border-transparent'
                                  }`}
                              >
                                <div className="flex items-center gap-2 overflow-hidden">
                                  <input
                                    type="checkbox"
                                    checked={isChecked}
                                    onChange={() => {
                                      if (isChecked) {
                                        setSelectedIRStudentIds(selectedIRStudentIds.filter(id => id !== student.id));
                                      } else {
                                        setSelectedIRStudentIds([...selectedIRStudentIds, student.id]);
                                      }
                                    }}
                                    className="w-4 h-4 text-indigo-600 rounded border-slate-300 focus:ring-indigo-500"
                                  />
                                  <span className="truncate">{student.name}</span>
                                </div>
                                <span className="text-[10px] text-slate-400 font-bold shrink-0 ml-1">Class {student.class}</span>
                              </label>
                            );
                          })
                      ) : (
                        <div className="text-[11px] text-slate-400 text-center py-3">No students found</div>
                      )}
                    </div>
                  </div>
                </div>

                {/* OK Button */}
                <div className="flex items-center justify-end gap-3 pt-2">
                  <button
                    type="button"
                    onClick={() => setShowIRDateModal(false)}
                    className="px-5 py-2.5 text-xs font-extrabold text-slate-600 bg-white border border-slate-200 rounded-xl hover:bg-slate-50 transition-all"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      if (selectedIRStudentIds.length === 0 && dateModalNextAction !== 'ADMIN_REPORT') {
                        alert('Please select at least one student.');
                        return;
                      }
                      setShowIRDateModal(false);
                      if (dateModalNextAction === 'ADMIN_REPORT') {
                        setShowDownloadModal(true);
                      } else {
                        setSelectedIRStudent(selectedIRStudentIds.length === irAssignedStudents.length ? 'ALL' : 'MULTI');
                        setShowIRModal(true);
                      }
                    }}
                    className="px-8 py-3 bg-[#1A365D] hover:bg-[#2A4365] text-white font-extrabold text-xs uppercase tracking-wider rounded-xl shadow-lg shadow-blue-900/20 transition-all flex items-center gap-2 active:scale-[0.98]"
                  >
                    <span>OK - Generate IR Report</span>
                    <ChevronRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Individual Report Modal */}
        {showIRModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[110] animate-fade-in" onClick={e => e.stopPropagation()}>
            <div className="bg-slate-100 rounded-2xl w-full max-w-5xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col h-[90vh]">
              <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-white shrink-0">
                <div className="flex items-center gap-2 text-[#1A365D]">
                  <FileText className="w-5 h-5 text-indigo-600" />
                  <h2 className="text-sm font-extrabold tracking-wider uppercase">Individual Report (IR)</h2>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={() => {
                      const csvLines = ["Date,Name,Class,Event Item,Reason / Notes"];

                      const studentsToExport = irAssignedStudents.filter(s => {
                        if (selectedIRStudent === 'ALL') return true;
                        if (selectedIRStudent === 'MULTI') return selectedIRStudentIds.includes(s.id);
                        if (selectedIRStudent && selectedIRStudent.id) return s.id === selectedIRStudent.id;
                        return selectedIRStudentIds.includes(s.id);
                      });

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
                            const day = d.getDate();
                            const month = d.getMonth() + 1;
                            const year = String(d.getFullYear()).slice(-2);
                            const formattedDate = `${day}/${month}/${year}`;

                            let eventDisplay = '';
                            const type = (log.event_type || '').toLowerCase();
                            const amt = Math.abs(Number(log.amount) || 1);
                            if (type === 'star') eventDisplay = `${amt} star`;
                            else if (type === 'tally') eventDisplay = `${amt} tally`;
                            else if (type === 'fine' || type === 'fin') eventDisplay = `fin`;
                            else if (type.includes('sheet') || type === 'yellow sheet') eventDisplay = `yello sheet`;
                            else if (type.includes('n&o') || type === 'neat') eventDisplay = `N&O tally`;
                            else if (type === 'spot fine' || type === 'spotfine' || type === 'room fine' || type === 'roomfine') eventDisplay = type.includes('room') ? `room fine` : `spot fine`;
                            else if (type === 'diary') eventDisplay = `${amt} diary`;
                            else eventDisplay = log.amount ? `${amt} ${log.event_type}` : log.event_type;

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
                      link.setAttribute("download", `Individual_Report_${irFromDate || 'all'}_to_${irToDate || 'all'}.csv`);
                      document.body.appendChild(link);
                      link.click();
                      document.body.removeChild(link);
                    }}
                    className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] uppercase font-bold rounded-lg shadow-sm transition-colors flex items-center gap-1"
                  >
                    <Download className="w-3.5 h-3.5" /> Download Excel
                  </button>
                  <button onClick={() => setShowIRModal(false)} className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors">
                    <X className="w-4 h-4" />
                  </button>
                </div>
              </div>

              <div className="p-6 overflow-y-auto flex-1 flex flex-col gap-8">
                {irAssignedStudents.filter(s => {
                  if (selectedIRStudent === 'ALL') return true;
                  if (selectedIRStudent === 'MULTI') return selectedIRStudentIds.includes(s.id);
                  if (selectedIRStudent && selectedIRStudent.id) return s.id === selectedIRStudent.id;
                  return selectedIRStudentIds.includes(s.id);
                }).map(student => {
                  const logsForStudent = irHistoryLogs
                    .filter(log => log.student_id === student.id)
                    .filter(log => {
                      const logDate = new Date(log.date).toISOString().split('T')[0];
                      return (!irFromDate || logDate >= irFromDate) && (!irToDate || logDate <= irToDate);
                    });

                  let starCount = 0;
                  let tallyCount = 0;
                  let fineSum = 0;
                  let noTallyCount = 0;
                  let yellowSheetCount = 0;
                  let diaryCount = 0;
                  let apologyCount = 0;

                  logsForStudent.forEach(log => {
                    const type = (log.event_type || '').toLowerCase();
                    const amt = Number(log.amount) || 1;
                    if (type === 'star') starCount += Math.abs(amt);
                    else if (type === 'tally') tallyCount += Math.abs(amt);
                    else if (type === 'fine' || type === 'fin' || type.includes('fine')) fineSum += Math.abs(amt);
                    else if (type.includes('n&o') || type === 'neat') noTallyCount += Math.abs(amt);
                    else if (type.includes('sheet') || type === 'yellow sheet') yellowSheetCount += Math.abs(amt);
                    else if (type === 'diary') diaryCount += Math.abs(amt);
                    else if (type === 'apology') apologyCount += Math.abs(amt);
                  });

                  return (
                    <div key={student.id} className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col gap-6">
                      {/* Header: Name, Class, Room */}
                      <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 border-b border-slate-100 pb-4">
                        <div className="flex items-center gap-4">
                          <div className="w-14 h-14 bg-indigo-600 text-white rounded-2xl font-black text-xl flex items-center justify-center shadow-md">
                            {getInitials(student.name)}
                          </div>
                          <div>
                            <h1 className="text-2xl font-black text-[#1A365D] uppercase tracking-tight">{student.name}</h1>
                            <div className="flex items-center gap-2 mt-1">
                              <span className="px-2.5 py-0.5 bg-indigo-50 text-indigo-700 rounded-md text-xs font-bold uppercase tracking-wide border border-indigo-100">
                                Class {student.class}
                              </span>
                              {student.roomNumber && (
                                <span className="px-2.5 py-0.5 bg-slate-100 text-slate-600 rounded-md text-xs font-bold uppercase tracking-wide">
                                  Room {student.roomNumber}
                                </span>
                              )}
                            </div>
                          </div>
                        </div>
                        <div className="text-right text-xs text-slate-400 font-bold">
                          Range: {irFromDate || 'Start'} to {irToDate || 'Present'}
                        </div>
                      </div>

                      {/* Stat Badges Grid */}
                      <div>
                        <h4 className="text-xs font-extrabold text-slate-500 uppercase tracking-wider mb-2">Summary Statistics</h4>
                        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2.5">
                          <div className="bg-emerald-50 border border-emerald-100 rounded-xl p-3 flex flex-col items-center justify-center text-center">
                            <div className="text-[10px] font-extrabold text-emerald-600 uppercase tracking-wider">Star</div>
                            <div className="text-xl font-black text-emerald-700 mt-0.5">{starCount || student.star || 0}</div>
                          </div>
                          <div className="bg-rose-50 border border-rose-100 rounded-xl p-3 flex flex-col items-center justify-center text-center">
                            <div className="text-[10px] font-extrabold text-rose-600 uppercase tracking-wider">Tally</div>
                            <div className="text-xl font-black text-rose-700 mt-0.5">{tallyCount || student.tally || 0}</div>
                          </div>
                          <div className="bg-red-50 border border-red-100 rounded-xl p-3 flex flex-col items-center justify-center text-center">
                            <div className="text-[10px] font-extrabold text-red-600 uppercase tracking-wider">Spot Fine</div>
                            <div className="text-xl font-black text-red-700 mt-0.5">₹{fineSum || student.fine || 0}</div>
                          </div>
                          <div className="bg-amber-50 border border-amber-100 rounded-xl p-3 flex flex-col items-center justify-center text-center">
                            <div className="text-[10px] font-extrabold text-amber-600 uppercase tracking-wider">N&O Tally</div>
                            <div className="text-xl font-black text-amber-700 mt-0.5">{noTallyCount || student.neatAndOrderTally || 0}</div>
                          </div>
                          <div className="bg-yellow-50 border border-yellow-200 rounded-xl p-3 flex flex-col items-center justify-center text-center">
                            <div className="text-[10px] font-extrabold text-yellow-700 uppercase tracking-wider">Yellow Sheet</div>
                            <div className="text-xl font-black text-yellow-800 mt-0.5">{yellowSheetCount || student.sheetTally || 0}</div>
                          </div>
                          <div className="bg-purple-50 border border-purple-100 rounded-xl p-3 flex flex-col items-center justify-center text-center">
                            <div className="text-[10px] font-extrabold text-purple-600 uppercase tracking-wider">Diary</div>
                            <div className="text-xl font-black text-purple-700 mt-0.5">{diaryCount || student.diaryTally || 0}</div>
                          </div>
                          <div className="bg-sky-50 border border-sky-100 rounded-xl p-3 flex flex-col items-center justify-center text-center">
                            <div className="text-[10px] font-extrabold text-sky-600 uppercase tracking-wider">Apology</div>
                            <div className="text-xl font-black text-sky-700 mt-0.5">{apologyCount || 0}</div>
                          </div>
                        </div>
                      </div>

                      {/* History Table: 3 Columns (Date | Event Item | Reason) */}
                      <div className="mt-1">
                        <h3 className="text-xs font-extrabold text-slate-800 mb-2 uppercase tracking-wider flex items-center gap-1.5">
                          <BarChart3 className="w-4 h-4 text-indigo-600" />
                          Report History
                        </h3>
                        <div className="overflow-x-auto rounded-xl border border-slate-200">
                          <table className="w-full text-left border-collapse text-xs">
                            <thead>
                              <tr className="bg-slate-100 border-b border-slate-200 text-[10px] font-extrabold text-slate-600 uppercase tracking-wider">
                                <th className="p-3 w-1/4">Date</th>
                                <th className="p-3 w-1/3">Event Item</th>
                                <th className="p-3 w-5/12">Reason / Notes</th>
                              </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-slate-100">
                              {logsForStudent.length > 0 ? (
                                logsForStudent.map(log => {
                                  const d = new Date(log.date);
                                  const day = d.getDate();
                                  const month = d.getMonth() + 1;
                                  const year = String(d.getFullYear()).slice(-2);
                                  const formattedDate = `${day}/${month}/${year}`;

                                  let eventDisplay = '';
                                  const type = (log.event_type || '').toLowerCase();
                                  const amt = Math.abs(Number(log.amount) || 1);
                                  if (type === 'star') eventDisplay = `${amt} star`;
                                  else if (type === 'tally') eventDisplay = `${amt} tally`;
                                  else if (type === 'fine' || type === 'fin') eventDisplay = `fin`;
                                  else if (type.includes('sheet') || type === 'yellow sheet') eventDisplay = `yello sheet`;
                                  else if (type.includes('n&o') || type === 'neat') eventDisplay = `N&O tally`;
                                  else if (type === 'spot fine' || type === 'spotfine' || type === 'room fine' || type === 'roomfine') eventDisplay = type.includes('room') ? `room fine` : `spot fine`;
                                  else if (type === 'diary') eventDisplay = `${amt} diary`;
                                  else eventDisplay = log.amount ? `${amt} ${log.event_type}` : log.event_type;

                                  return (
                                    <tr key={log.id} className="hover:bg-slate-50/80 transition-colors">
                                      <td className="p-3 font-semibold text-slate-700">{formattedDate}</td>
                                      <td className="p-3 font-extrabold text-[#1A365D] capitalize">{eventDisplay}</td>
                                      <td className="p-3 text-slate-600 font-medium">{log.reason || '-'}</td>
                                    </tr>
                                  );
                                })
                              ) : (
                                <tr>
                                  <td colSpan="3" className="p-6 text-center text-slate-400 font-medium italic">
                                    No history records found for this date range.
                                  </td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  );
                })}
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
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 z-[120] animate-fade-in overflow-y-auto">
            <div className="bg-white rounded-3xl w-[95%] sm:w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-4 sm:p-6 gap-4 sm:gap-5 max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2 text-emerald-700">
                  <Award className="w-6 h-6 shrink-0" />
                  <h3 className="font-extrabold text-base sm:text-lg">Room Tally Entry</h3>
                </div>
                <button onClick={() => setShowRoomTallyModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold p-1">✕</button>
              </div>

              <div className="flex flex-col gap-4">
                <div>
                  <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">Tally Number (Count)</label>
                  <div className="flex items-center gap-2 sm:gap-3">
                    <button
                      type="button"
                      onClick={() => setRoomTallyCount(prev => Math.max(1, (Number(prev) || 1) - 1))}
                      className="w-11 h-11 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-black text-lg flex items-center justify-center transition-colors shrink-0 active:scale-95 cursor-pointer"
                    >
                      -
                    </button>
                    <input
                      type="number"
                      inputMode="numeric"
                      pattern="[0-9]*"
                      min="1"
                      value={roomTallyCount}
                      onChange={e => {
                        const val = e.target.value;
                        if (val === '') {
                          setRoomTallyCount('');
                        } else {
                          const parsed = parseInt(val, 10);
                          setRoomTallyCount(isNaN(parsed) ? '' : Math.max(1, parsed));
                        }
                      }}
                      onBlur={() => {
                        if (roomTallyCount === '' || Number(roomTallyCount) < 1) {
                          setRoomTallyCount(3);
                        }
                      }}
                      placeholder="3"
                      className="flex-1 py-2.5 px-3 sm:px-4 bg-slate-50 border border-slate-200 rounded-xl font-black text-center text-xl text-slate-800 focus:outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-500/20"
                    />
                    <button
                      type="button"
                      onClick={() => setRoomTallyCount(prev => (Number(prev) || 0) + 1)}
                      className="w-11 h-11 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-black text-lg flex items-center justify-center transition-colors shadow-xs shrink-0 active:scale-95 cursor-pointer"
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
                  type="button"
                  onClick={() => setShowRoomTallyModal(false)}
                  className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const currentRoomKey = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
                    const allocated = getHostelRoomStudents(selectedHostel, currentRoomKey);
                    const selectedObjs = allocated.filter(s => performanceSelectedStudents.includes(s.id));

                    const amount = Math.max(1, Number(roomTallyCount) || 3);
                    const reason = roomTallyReason || 'Room Tally';

                    let studentsToUpsert = [];
                    let updatedStudentsList = [...students];

                    selectedObjs.forEach(so => {
                      const matchIndex = updatedStudentsList.findIndex(s =>
                        s.id === so.id ||
                        (s.name.trim().toLowerCase() === so.name.trim().toLowerCase() && (s.class || '').trim().toLowerCase() === (so.class || '').trim().toLowerCase())
                      );

                      if (matchIndex !== -1) {
                        const s = updatedStudentsList[matchIndex];
                        const updated = {
                          ...s,
                          tally: (s.tally || 0) + amount,
                          neatAndOrderTally: (s.neatAndOrderTally || 0) + amount,
                          neatAndOrderIncidents: (s.neatAndOrderIncidents || 0) + 1,
                          neatAndOrderReason: reason || s.neatAndOrderReason
                        };
                        logHistory(s.id, 'N&O Tally', amount, reason);
                        studentsToUpsert.push(updated);
                        updatedStudentsList[matchIndex] = updated;
                      } else {
                        const newId = so.id && !so.id.startsWith('r') ? so.id : `st-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
                        const newStudent = {
                          id: newId,
                          name: so.name,
                          class: (so.class || 'S1B').toUpperCase(),
                          star: 0,
                          tally: amount,
                          starReason: '',
                          tallyReason: '',
                          diaryStar: 0,
                          diaryTally: 0,
                          neatAndOrderTally: amount,
                          neatAndOrderReason: reason,
                          neatAndOrderIncidents: 1,
                          fine: 0,
                          fineCount: 0,
                          fineReason: '',
                          room: currentRoomKey,
                          roomNumber: currentRoomKey,
                          hostelBlock: selectedHostel
                        };
                        logHistory(newId, 'N&O Tally', amount, reason);
                        studentsToUpsert.push(newStudent);
                        updatedStudentsList.push(newStudent);
                      }
                    });

                    if (studentsToUpsert.length > 0) {
                      setStudents(updatedStudentsList);
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
                    const allocated = getHostelRoomStudents(selectedHostel, currentRoomKey);
                    const selectedObjs = allocated.filter(s => performanceSelectedStudents.includes(s.id));

                    const fineAmt = Number(roomFineAmount) || 0;
                    const reason = roomFineReason || 'Room Fine Violation';

                    let studentsToUpsert = [];
                    let updatedStudentsList = [...students];

                    selectedObjs.forEach(so => {
                      const matchIndex = updatedStudentsList.findIndex(s =>
                        s.id === so.id ||
                        (s.name.trim().toLowerCase() === so.name.trim().toLowerCase() && (s.class || '').trim().toLowerCase() === (so.class || '').trim().toLowerCase())
                      );

                      if (matchIndex !== -1) {
                        const s = updatedStudentsList[matchIndex];
                        const updated = {
                          ...s,
                          fine: (s.fine || 0) + fineAmt,
                          fineCount: (s.fineCount || 0) + 1,
                          fineReason: reason || s.fineReason,
                          spotFine: (s.spotFine || 0) + fineAmt,
                          spotFineReason: reason || s.spotFineReason
                        };
                        logHistory(s.id, 'Room Fine', fineAmt, reason);
                        studentsToUpsert.push(updated);
                        updatedStudentsList[matchIndex] = updated;
                      } else {
                        const newId = so.id && !so.id.startsWith('r') ? so.id : `st-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
                        const newStudent = {
                          id: newId,
                          name: so.name,
                          class: (so.class || 'S1B').toUpperCase(),
                          star: 0,
                          tally: 0,
                          starReason: '',
                          tallyReason: '',
                          diaryStar: 0,
                          diaryTally: 0,
                          neatAndOrderTally: 0,
                          neatAndOrderReason: '',
                          neatAndOrderIncidents: 0,
                          fine: fineAmt,
                          fineCount: 1,
                          fineReason: reason,
                          spotFine: fineAmt,
                          spotFineReason: reason,
                          room: currentRoomKey,
                          roomNumber: currentRoomKey,
                          hostelBlock: selectedHostel
                        };
                        logHistory(newId, 'Room Fine', fineAmt, reason);
                        studentsToUpsert.push(newStudent);
                        updatedStudentsList.push(newStudent);
                      }
                    });

                    if (studentsToUpsert.length > 0) {
                      setStudents(updatedStudentsList);
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

        {/* ADD STUDENT TO ROOM MODAL */}
        {showAddRoomStudentModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[120] animate-fade-in">
            <div className="bg-white rounded-3xl w-full max-w-lg shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-6 gap-4" onClick={e => e.stopPropagation()}>

              {/* Modal Header */}
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2 text-emerald-600">
                  <PlusCircle className="w-6 h-6" />
                  <h3 className="font-extrabold text-lg">Add Student to {selectedRoom}</h3>
                </div>
                <button onClick={() => setShowAddRoomStudentModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
              </div>

              {/* Modal Mode Tabs: Single vs Bulk / Excel */}
              <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200">
                <button
                  type="button"
                  onClick={() => setAddRoomStudentModalTab('single')}
                  className={`flex-1 py-2 rounded-lg text-xs font-black uppercase tracking-wider transition-all flex items-center justify-center gap-1.5 ${addRoomStudentModalTab === 'single'
                      ? 'bg-white text-emerald-700 shadow-xs'
                      : 'text-slate-500 hover:text-slate-800'
                    }`}
                >
                  <Plus className="w-4 h-4" />
                  <span>Single Student</span>
                </button>
                <button
                  type="button"
                  onClick={() => setAddRoomStudentModalTab('bulk')}
                  className={`flex-1 py-2 rounded-lg text-xs font-black uppercase tracking-wider transition-all flex items-center justify-center gap-1.5 ${addRoomStudentModalTab === 'bulk'
                      ? 'bg-emerald-600 text-white shadow-xs'
                      : 'text-slate-500 hover:text-slate-800'
                    }`}
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  <span>Excel / Bulk Import</span>
                </button>
              </div>

              {addRoomStudentModalTab === 'single' ? (
                /* SINGLE ADD FORM */
                <form onSubmit={handleAddStudentToPerformanceRoom} className="flex flex-col gap-4">
                  {/* Select Existing Student or Type Name */}
                  <div className="relative">
                    <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">Student Name</label>
                    <div className="relative">
                      <Search className="w-5 h-5 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        value={addRoomStudentName}
                        onChange={e => {
                          setAddRoomStudentName(e.target.value);
                          setShowAddRoomDropdown(true);
                        }}
                        onFocus={() => setShowAddRoomDropdown(true)}
                        placeholder="Search or enter student name..."
                        className="w-full py-3 pl-11 pr-4 bg-slate-50 border border-slate-200 rounded-xl font-bold text-sm text-slate-800 focus:outline-none focus:border-emerald-500"
                        required
                      />
                    </div>

                    {/* Dropdown suggestions from existing students */}
                    {showAddRoomDropdown && addRoomStudentName && (
                      <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-slate-200 rounded-xl shadow-lg max-h-48 overflow-y-auto z-50 divide-y divide-slate-100">
                        {students
                          .filter(s => s.name.toLowerCase().includes(addRoomStudentName.toLowerCase()))
                          .slice(0, 10)
                          .map(s => (
                            <button
                              key={s.id}
                              type="button"
                              onMouseDown={() => {
                                setAddRoomStudentName(s.name);
                                if (s.class) setAddRoomStudentClass(s.class);
                                setShowAddRoomDropdown(false);
                              }}
                              className="w-full text-left px-4 py-2.5 text-xs text-slate-700 hover:bg-emerald-50 hover:text-emerald-800 flex justify-between items-center transition-colors"
                            >
                              <span className="font-extrabold">{s.name}</span>
                              <span className="font-bold text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded-md uppercase">Class {s.class}</span>
                            </button>
                          ))}
                      </div>
                    )}
                  </div>

                  {/* Class Selection */}
                  <div>
                    <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">Class</label>
                    <select
                      value={addRoomStudentClass}
                      onChange={e => setAddRoomStudentClass(e.target.value)}
                      className="w-full py-3 px-4 bg-slate-50 border border-slate-200 rounded-xl font-bold text-sm text-slate-800 focus:outline-none focus:border-emerald-500 uppercase"
                      required
                    >
                      {CLASSES.map(cls => (
                        <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>
                      ))}
                    </select>
                  </div>

                  <div className="flex items-center gap-3 pt-3">
                    <button
                      type="button"
                      onClick={() => setShowAddRoomStudentModal(false)}
                      className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors"
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-1.5"
                    >
                      <Plus className="w-4 h-4" /> Add Student
                    </button>
                  </div>
                </form>
              ) : (
                /* BULK EXCEL IMPORT FORM */
                <form onSubmit={handleBulkImportStudentsToRoom} className="flex flex-col gap-4">

                  {/* Excel File Upload Box */}
                  <div>
                    <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">
                      Option 1: Upload Excel File (.xlsx, .xls, .csv)
                    </label>
                    <input
                      type="file"
                      accept=".xlsx, .xls, .csv"
                      id="bulk-excel-file-input"
                      onChange={handleBulkExcelUpload}
                      className="hidden"
                    />
                    <label
                      htmlFor="bulk-excel-file-input"
                      className="w-full py-3 px-4 bg-emerald-50 border-2 border-dashed border-emerald-300 hover:border-emerald-500 rounded-xl flex items-center justify-center gap-2 text-emerald-800 font-bold text-xs cursor-pointer hover:bg-emerald-100 transition-all"
                    >
                      <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
                      <span>{bulkImportFileName ? `File Selected: ${bulkImportFileName}` : 'Choose Excel / CSV File'}</span>
                    </label>
                  </div>

                  {/* Textarea for Pasting Data */}
                  <div>
                    <div className="flex items-center justify-between mb-1.5">
                      <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider">
                        Option 2: Paste Student List
                      </label>
                      <button
                        type="button"
                        onClick={() => {
                          const sample = `Nitha Fathima  (C2A)\nRifa Fathima P  (C2A)\nSana Fathima VK (S2A)\nUmmu habeeba M (C2A)`;
                          handleBulkTextChange(sample);
                        }}
                        className="text-[11px] text-emerald-600 hover:text-emerald-800 font-extrabold underline"
                      >
                        Insert Sample Format
                      </button>
                    </div>
                    <textarea
                      rows={4}
                      value={bulkImportText}
                      onChange={e => handleBulkTextChange(e.target.value)}
                      placeholder={`Format:\nNitha Fathima (C2A)\nRifa Fathima P (C2A)\nSana Fathima VK (S2A)\nUmmu habeeba M (C2A)`}
                      className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-mono text-xs text-slate-800 focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  {/* Parsed Preview List */}
                  {bulkParsedStudents.length > 0 && (
                    <div className="flex flex-col gap-2 bg-slate-50 p-3 rounded-2xl border border-slate-200">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-extrabold text-emerald-700 uppercase tracking-wider">
                          Parsed Preview ({bulkParsedStudents.length} Students)
                        </span>
                        <button
                          type="button"
                          onClick={() => {
                            setBulkParsedStudents([]);
                            setBulkImportText('');
                            setBulkImportFileName('');
                          }}
                          className="text-[10px] font-bold text-rose-600 hover:underline"
                        >
                          Clear All
                        </button>
                      </div>

                      <div className="max-h-36 overflow-y-auto divide-y divide-slate-200 border border-slate-200 rounded-xl bg-white">
                        {bulkParsedStudents.map((s, idx) => (
                          <div key={idx} className="p-2 flex items-center justify-between text-xs gap-2">
                            <div className="flex items-center gap-2 overflow-hidden">
                              <span className="text-[10px] font-mono font-bold text-slate-400 w-5">#{idx + 1}</span>
                              <span className="font-extrabold text-slate-800 truncate">{s.name}</span>
                            </div>
                            <div className="flex items-center gap-2">
                              <span className="font-extrabold text-[10px] bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-md uppercase">
                                {s.class || 'S1B'}
                              </span>
                              <button
                                type="button"
                                onClick={() => {
                                  const updated = bulkParsedStudents.filter((_, i) => i !== idx);
                                  setBulkParsedStudents(updated);
                                  setBulkImportText(updated.map(x => `${x.name} (${x.class})`).join('\n'));
                                }}
                                className="text-slate-400 hover:text-rose-600 font-bold"
                              >
                                ✕
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="flex items-center gap-3 pt-2">
                    <button
                      type="button"
                      onClick={() => setShowAddRoomStudentModal(false)}
                      className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors"
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      disabled={bulkParsedStudents.length === 0}
                      className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-1.5"
                    >
                      <Plus className="w-4 h-4" /> Import {bulkParsedStudents.length > 0 ? `(${bulkParsedStudents.length})` : ''} Students
                    </button>
                  </div>
                </form>
              )}

            </div>
          </div>
        )}

        {/* REMOVE STUDENT FROM ROOM MODAL */}
        {showRemoveRoomStudentModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[120] animate-fade-in">
            <div className="bg-white rounded-3xl w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-6 gap-5" onClick={e => e.stopPropagation()}>
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2 text-rose-600">
                  <Trash2 className="w-6 h-6" />
                  <h3 className="font-extrabold text-lg">Remove Student from {selectedRoom}</h3>
                </div>
                <button onClick={() => setShowRemoveRoomStudentModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
              </div>

              {(() => {
                const currentRoomKey = selectedRoom ? String(selectedRoom).replace(/room\s*/i, '').trim() : '';
                const roomStudents = roomStudentMapping[currentRoomKey] !== undefined ?
                  roomStudentMapping[currentRoomKey] :
                  students.filter(s => String(s.room || s.roomNumber || '').trim() === currentRoomKey);

                return (
                  <div className="flex flex-col gap-4">
                    <p className="text-xs font-semibold text-slate-500">
                      Select the student(s) you wish to remove from <span className="font-extrabold text-slate-800">{selectedRoom}</span>:
                    </p>

                    <div className="max-h-60 overflow-y-auto space-y-2 pr-1">
                      {roomStudents.map(s => {
                        const isSelected = removeSelectedRoomStudentIds.includes(s.id);
                        return (
                          <div
                            key={s.id}
                            onClick={() => {
                              setRemoveSelectedRoomStudentIds(prev =>
                                prev.includes(s.id) ? prev.filter(id => id !== s.id) : [...prev, s.id]
                              );
                            }}
                            className={`p-3 rounded-2xl border flex items-center justify-between cursor-pointer transition-all ${isSelected
                                ? 'bg-rose-50 border-rose-400 text-rose-950'
                                : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-slate-300'
                              }`}
                          >
                            <div className="flex items-center gap-2.5">
                              <div className={`w-7 h-7 rounded-lg flex items-center justify-center font-bold text-xs ${isSelected ? 'bg-rose-600 text-white' : 'bg-slate-200 text-slate-600'
                                }`}>
                                {s.name.charAt(0)}
                              </div>
                              <div>
                                <span className="font-extrabold text-xs block">{s.name}</span>
                                <span className="text-[10px] font-bold text-slate-400 uppercase">Class {s.class}</span>
                              </div>
                            </div>

                            <div className={`w-5 h-5 rounded-md border-2 flex items-center justify-center ${isSelected ? 'bg-rose-600 border-rose-600' : 'border-slate-300 bg-white'
                              }`}>
                              {isSelected && <Check className="w-3.5 h-3.5 text-white stroke-[3]" />}
                            </div>
                          </div>
                        );
                      })}

                      {roomStudents.length === 0 && (
                        <div className="p-6 text-center text-slate-400 font-bold text-xs">
                          No students currently allocated in this room.
                        </div>
                      )}
                    </div>

                    <div className="flex items-center gap-3 pt-3 border-t border-slate-100">
                      <button
                        type="button"
                        onClick={() => setShowRemoveRoomStudentModal(false)}
                        className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors"
                      >
                        Cancel
                      </button>
                      <button
                        type="button"
                        disabled={removeSelectedRoomStudentIds.length === 0}
                        onClick={() => handleRemoveStudentsFromPerformanceRoom(removeSelectedRoomStudentIds)}
                        className={`flex-1 py-3 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all flex items-center justify-center gap-1.5 ${removeSelectedRoomStudentIds.length === 0
                            ? 'bg-slate-300 cursor-not-allowed shadow-none'
                            : 'bg-rose-600 hover:bg-rose-700 active:scale-95'
                          }`}
                      >
                        <Trash2 className="w-4 h-4" /> Remove ({removeSelectedRoomStudentIds.length})
                      </button>
                    </div>
                  </div>
                );
              })()}
            </div>
          </div>
        )}

        {/* ADD NEW ROOM MODAL */}
        {showAddNewRoomModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[120] animate-fade-in">
            <div className="bg-white rounded-3xl w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-6 gap-5" onClick={e => e.stopPropagation()}>
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex flex-col gap-1">
                  <span className="text-[10px] font-black text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md uppercase tracking-wider self-start">
                    {activeHostelName}
                  </span>
                  <div className="flex items-center gap-2 text-emerald-600">
                    <PlusCircle className="w-5 h-5" />
                    <h3 className="font-extrabold text-lg">Add New Room</h3>
                  </div>
                </div>
                <button onClick={() => setShowAddNewRoomModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
              </div>

              <form onSubmit={handleAddNewRoom} className="flex flex-col gap-4">
                <div>
                  <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">Room Number / Identifier</label>
                  <input
                    type="text"
                    value={newRoomNumberInput}
                    onChange={e => setNewRoomNumberInput(e.target.value)}
                    placeholder="e.g. 236, 101A, B-12..."
                    className="w-full py-3 px-4 bg-slate-50 border border-slate-200 rounded-xl font-bold text-sm text-slate-800 focus:outline-none focus:border-emerald-500"
                    required
                    autoFocus
                  />
                </div>

                <div className="flex items-center gap-3 pt-3">
                  <button
                    type="button"
                    onClick={() => setShowAddNewRoomModal(false)}
                    className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-1.5"
                  >
                    <Plus className="w-4 h-4" /> Add Room
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* DELETE ROOM MODAL */}
        {showDeleteRoomModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[120] animate-fade-in">
            <div className="bg-white rounded-3xl w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-6 gap-5" onClick={e => e.stopPropagation()}>
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex flex-col gap-1">
                  <span className="text-[10px] font-black text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded-md uppercase tracking-wider self-start">
                    {activeHostelName}
                  </span>
                  <div className="flex items-center gap-2 text-rose-600">
                    <Trash2 className="w-5 h-5" />
                    <h3 className="font-extrabold text-lg">Delete Room(s)</h3>
                  </div>
                </div>
                <button onClick={() => setShowDeleteRoomModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
              </div>


              <div className="flex flex-col gap-4">
                <p className="text-xs font-semibold text-slate-500">
                  Select room(s) you wish to remove from the room list:
                </p>

                <div className="max-h-60 overflow-y-auto grid grid-cols-3 gap-2 pr-1">
                  {hostelRoomList.map(num => {
                    const isSelected = roomsToDelete.includes(num);
                    return (
                      <button
                        key={num}
                        type="button"
                        onClick={() => {
                          setRoomsToDelete(prev =>
                            prev.includes(num) ? prev.filter(n => n !== num) : [...prev, num]
                          );
                        }}
                        className={`p-3 rounded-xl border font-extrabold text-xs flex items-center justify-between transition-all ${isSelected
                            ? 'bg-rose-50 border-rose-400 text-rose-950 shadow-xs'
                            : 'bg-slate-50 border-slate-200 text-slate-700 hover:border-slate-300'
                          }`}
                      >
                        <span>Room {num}</span>
                        {isSelected && <Check className="w-4 h-4 text-rose-600" />}
                      </button>
                    );
                  })}
                </div>

                <div className="flex items-center gap-3 pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setShowDeleteRoomModal(false)}
                    className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    disabled={roomsToDelete.length === 0}
                    onClick={handleDeleteRooms}
                    className={`flex-1 py-3 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all flex items-center justify-center gap-1.5 ${roomsToDelete.length === 0
                        ? 'bg-slate-300 cursor-not-allowed shadow-none'
                        : 'bg-rose-600 hover:bg-rose-700 active:scale-95'
                      }`}
                  >
                    <Trash2 className="w-4 h-4" /> Delete ({roomsToDelete.length})
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* EDIT ROOM NUMBER MODAL (LONG PRESS) */}
        {showEditRoomNumberModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[120] animate-fade-in">
            <div className="bg-white rounded-3xl w-full max-w-md shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-6 gap-5" onClick={e => e.stopPropagation()}>
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2 text-indigo-600">
                  <Pencil className="w-6 h-6" />
                  <h3 className="font-extrabold text-lg">Change Room Number</h3>
                </div>
                <button onClick={() => setShowEditRoomNumberModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
              </div>

              <form onSubmit={handleEditRoomNumberSubmit} className="flex flex-col gap-4">
                <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl text-xs">
                  <span className="font-semibold text-slate-500">Current Room: </span>
                  <span className="font-extrabold text-slate-800">Room {editingRoomOldNum}</span>
                </div>

                <div>
                  <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">New Room Number / Identifier</label>
                  <input
                    type="text"
                    value={editingRoomNewNum}
                    onChange={e => setEditingRoomNewNum(e.target.value)}
                    className="w-full py-3 px-4 bg-slate-50 border border-slate-200 rounded-xl font-bold text-sm text-slate-800 focus:outline-none focus:border-indigo-500"
                    required
                    autoFocus
                  />
                </div>

                <div className="flex items-center gap-3 pt-3">
                  <button
                    type="button"
                    onClick={() => setShowEditRoomNumberModal(false)}
                    className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-colors"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="flex-1 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-1.5"
                  >
                    <Pencil className="w-4 h-4" /> Save Number
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* All Activities Report Modal */}
        {showActivitiesReportModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 z-[150] animate-fade-in">
            <div className="bg-white rounded-2xl w-full max-w-4xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col h-[90vh]" onClick={e => e.stopPropagation()}>

              {/* Modal Header */}
              <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-[#1A365D] text-white shrink-0">
                <div className="flex items-center gap-2.5">
                  <FileText className="w-5 h-5 text-emerald-400" />
                  <div>
                    <h2 className="text-base font-black tracking-wider uppercase">All Activities Report</h2>
                    <p className="text-xs text-slate-300 font-medium">View, Filter, Edit & Delete Activity Records</p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  {isAdminAuthenticated && (
                    <button
                      onClick={() => {
                        setNewActivityForm({
                          student_id: students[0]?.id || '',
                          event_type: 'tally',
                          amount: 1,
                          reason: '',
                          date: new Date().toISOString().split('T')[0]
                        });
                        setShowAddActivityModal(true);
                      }}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition-all shadow-xs cursor-pointer"
                    >
                      <Plus className="w-4 h-4" />
                      <span>Add Record</span>
                    </button>
                  )}
                  <button
                    onClick={() => setShowActivitiesReportModal(false)}
                    className="text-slate-300 hover:text-white hover:bg-white/10 p-1.5 rounded-lg transition-colors cursor-pointer"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </div>

              {/* Filter Bar */}
              <div className="p-4 bg-slate-50 border-b border-slate-200 grid grid-cols-1 sm:grid-cols-3 gap-3 shrink-0">
                {/* Date Filter */}
                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase tracking-wider mb-1 flex justify-between items-center">
                    <span>Select Date</span>
                    {activityDateFilter && (
                      <button
                        type="button"
                        onClick={() => setActivityDateFilter('')}
                        className="text-[10px] text-blue-600 hover:underline capitalize cursor-pointer font-extrabold"
                      >
                        Show All Dates
                      </button>
                    )}
                  </label>
                  <input
                    type="date"
                    value={activityDateFilter}
                    onChange={(e) => setActivityDateFilter(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-[#1A365D]"
                  />
                </div>

                {/* Class Filter */}
                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase tracking-wider mb-1">
                    Select Class
                  </label>
                  <select
                    value={activityClassFilter}
                    onChange={(e) => setActivityClassFilter(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-[#1A365D] uppercase"
                  >
                    <option value="ALL">All Classes</option>
                    {CLASSES.map(cls => (
                      <option key={cls} value={cls}>Class {cls.toUpperCase()}</option>
                    ))}
                  </select>
                </div>

                {/* Search Query */}
                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase tracking-wider mb-1">
                    Search Student / Reason
                  </label>
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      placeholder="Type name or reason..."
                      value={activitySearchQuery}
                      onChange={(e) => setActivitySearchQuery(e.target.value)}
                      className="w-full pl-8 pr-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-700 focus:outline-none focus:border-[#1A365D]"
                    />
                  </div>
                </div>
              </div>

              {/* Selection & Action Toolbar Bar */}
              {isAdminAuthenticated && (
                <div className="px-4 py-2.5 bg-slate-100 border-b border-slate-200 flex items-center justify-between gap-2 shrink-0">
                  <div className="flex items-center gap-2">
                    <input
                      type="checkbox"
                      id="selectAllActivities"
                      checked={filteredActivities.length > 0 && selectedActivityIds.length === filteredActivities.length}
                      onChange={(e) => {
                        if (e.target.checked) {
                          setSelectedActivityIds(filteredActivities.map(a => a.id));
                        } else {
                          setSelectedActivityIds([]);
                        }
                      }}
                      className="w-4 h-4 rounded border-slate-300 text-[#1A365D] focus:ring-blue-500 cursor-pointer"
                    />
                    <label htmlFor="selectAllActivities" className="text-xs font-extrabold text-slate-700 cursor-pointer">
                      Select All ({filteredActivities.length})
                    </label>
                  </div>

                  {selectedActivityIds.length > 0 && (
                    <button
                      onClick={handleBulkDeleteActivities}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold rounded-lg shadow-xs active:scale-95 transition-all cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>Delete Selected ({selectedActivityIds.length})</span>
                    </button>
                  )}
                </div>
              )}

              {/* Activities Table / Card List */}
              <div className="flex-1 overflow-y-auto p-3 sm:p-4 bg-slate-50/50">
                {filteredActivities.length > 0 ? (
                  <div className="flex flex-col gap-2.5">
                    {filteredActivities.map((log) => {
                      const studentObj = students.find(s => String(s.id) === String(log.student_id));
                      const studentName = studentObj ? studentObj.name : (log.student_name || `Student #${log.student_id}`);
                      const studentClass = studentObj ? studentObj.class.toUpperCase() : (log.class ? log.class.toUpperCase() : '-');

                      const d = log.date ? new Date(log.date) : new Date();
                      const formattedDate = isNaN(d.getTime()) ? String(log.date) : `${d.getDate()}/${d.getMonth() + 1}/${String(d.getFullYear()).slice(-2)}`;

                      let eventDisplay = '';
                      const type = String(log.event_type || '').toLowerCase();
                      if (type === 'star') {
                        eventDisplay = `${log.amount || 1} STAR`;
                      } else if (type === 'tally') {
                        eventDisplay = `${log.amount || 1} TALLY`;
                      } else if (type === 'spot fine' || type === 'spot_fine' || type === 'room fine' || type === 'room_fine' || type === 'fine') {
                        const fineLabel = type.includes('room') ? 'ROOM FINE' : 'SPOT FINE';
                        eventDisplay = `${fineLabel} ${log.amount ? '(₹' + log.amount + ')' : ''}`;
                      } else if (type === 'spot') {
                        eventDisplay = `SPOT`;
                      } else if (type === 'sata') {
                        eventDisplay = `${log.amount || 1} SATA`;
                      } else {
                        eventDisplay = `${log.amount && log.amount > 0 ? log.amount + ' ' : ''}${type.toUpperCase()}`;
                      }

                      const isChecked = selectedActivityIds.includes(log.id);

                      return (
                        <div
                          key={log.id}
                          className={`p-3.5 bg-white border rounded-xl shadow-xs transition-all flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${isChecked ? 'border-blue-400 bg-blue-50/30' : 'border-slate-200 hover:border-slate-300'
                            }`}
                        >
                          <div className="flex items-start sm:items-center gap-3 min-w-0 flex-1">
                            {isAdminAuthenticated && (
                              <input
                                type="checkbox"
                                checked={isChecked}
                                onChange={() => {
                                  setSelectedActivityIds(prev =>
                                    prev.includes(log.id) ? prev.filter(i => i !== log.id) : [...prev, log.id]
                                  );
                                }}
                                className="w-4 h-4 mt-0.5 sm:mt-0 rounded border-slate-300 text-[#1A365D] focus:ring-blue-500 cursor-pointer shrink-0"
                              />
                            )}
                            <div className="min-w-0 flex-1 flex flex-col sm:flex-row sm:items-center gap-1.5 sm:gap-4">
                              <span className="font-mono text-[11px] font-bold text-slate-500 shrink-0">
                                {formattedDate}
                              </span>
                              <span className="font-extrabold text-xs text-slate-900 uppercase truncate">
                                {studentName}
                              </span>
                              <span className="inline-block self-start sm:self-auto px-2 py-0.5 rounded-md bg-blue-50 border border-blue-200 text-[10px] font-extrabold text-blue-800 shrink-0">
                                {studentClass}
                              </span>
                              <span className={`inline-block self-start sm:self-auto px-2.5 py-0.5 rounded-lg text-[10px] font-black shrink-0 ${type.includes('star') || type.includes('sata') ? 'bg-amber-100 text-amber-800 border border-amber-200' :
                                  type.includes('fine') || type.includes('spot') ? 'bg-rose-100 text-rose-800 border border-rose-200' :
                                    type.includes('tally') ? 'bg-sky-100 text-sky-800 border border-sky-200' :
                                      'bg-purple-100 text-purple-800 border border-purple-200'
                                }`}>
                                {eventDisplay}
                              </span>
                              {log.reason && (
                                <span className="text-xs text-slate-600 font-medium italic truncate">
                                  "{log.reason}"
                                </span>
                              )}
                            </div>
                          </div>

                          {isAdminAuthenticated && (
                            <div className="flex items-center gap-1.5 self-end sm:self-auto shrink-0 border-t sm:border-t-0 pt-2 sm:pt-0 w-full sm:w-auto justify-end">
                              <button
                                onClick={() => setActivityEditModal({ ...log, date: log.date ? new Date(log.date).toISOString().split('T')[0] : new Date().toISOString().split('T')[0] })}
                                className="px-2.5 py-1.5 rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-600 text-xs font-bold transition-colors flex items-center gap-1 cursor-pointer"
                                title="Edit Activity"
                              >
                                <Pencil className="w-3.5 h-3.5" />
                                <span>Edit</span>
                              </button>
                              <button
                                onClick={() => handleDeleteActivity(log.id)}
                                className="px-2.5 py-1.5 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-600 text-xs font-bold transition-colors flex items-center gap-1 cursor-pointer"
                                title="Delete Activity"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                                <span>Delete</span>
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                ) : (
                  <div className="p-8 text-center text-slate-400 font-medium italic bg-white rounded-xl border border-slate-200">
                    No activities found matching your filter criteria.
                  </div>
                )}
              </div>

              {/* Modal Footer */}
              <div className="p-4 bg-white border-t border-slate-200 flex justify-between items-center text-xs font-semibold text-slate-500 shrink-0">
                <span>Showing {filteredActivities.length} activity records ({selectedActivityIds.length} selected)</span>
                <button
                  onClick={() => setShowActivitiesReportModal(false)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl transition-all cursor-pointer"
                >
                  Close
                </button>
              </div>

            </div>
          </div>
        )}


        {/* Edit Activity Modal */}
        {/* Change Hostel Name Modal */}
        {showEditHostelModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[200] animate-fade-in">
            <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl border border-slate-200 overflow-hidden flex flex-col scale-in">
              <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-[#1A365D] text-white">
                <div className="flex items-center gap-2">
                  <Pencil className="w-4 h-4 text-amber-400" />
                  <h3 className="font-extrabold tracking-wider text-sm uppercase">Change Hostel Name</h3>
                </div>
                <button onClick={() => setShowEditHostelModal(false)} className="text-slate-300 hover:text-white cursor-pointer">
                  <X className="w-4 h-4" />
                </button>
              </div>
              <form onSubmit={handleEditHostelNameSubmit} className="p-5 flex flex-col gap-4 text-xs font-semibold">
                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Hostel Block Name</label>
                  <input
                    type="text"
                    value={editingHostelNewName}
                    onChange={(e) => setEditingHostelNewName(e.target.value)}
                    placeholder="E.g. COVE, BOYS WAVES, HIVE..."
                    className="w-full p-3.5 border border-slate-200 rounded-xl font-extrabold text-slate-800 bg-slate-50 focus:outline-none focus:border-[#1A365D]"
                    required
                    autoFocus
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-2">Quick Presets</label>
                  <div className="flex flex-wrap gap-1.5">
                    {['MAIN BLOCK BOYS', 'BOYS WAVES', 'COVE', 'HIVE', 'BLOCK A', 'BLOCK B'].map(preset => (
                      <button
                        key={preset}
                        type="button"
                        onClick={() => setEditingHostelNewName(preset)}
                        className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition-all border cursor-pointer ${editingHostelNewName.trim().toUpperCase() === preset
                            ? 'bg-[#1A365D] text-white border-[#1A365D]'
                            : 'bg-slate-100 text-slate-600 border-slate-200 hover:bg-slate-200'
                          }`}
                      >
                        {preset}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="flex gap-2 justify-end pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setShowEditHostelModal(false)}
                    className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold rounded-xl cursor-pointer transition-colors"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2.5 bg-[#1A365D] hover:bg-blue-900 text-white font-extrabold rounded-xl shadow-md cursor-pointer transition-all active:scale-95"
                  >
                    Save Changes
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {activityEditModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[200] animate-fade-in">
            <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl border border-slate-200 overflow-hidden flex flex-col scale-in">
              <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-[#1A365D] text-white">
                <div className="flex items-center gap-2">
                  <Pencil className="w-4 h-4 text-amber-400" />
                  <h3 className="font-extrabold tracking-wider text-sm uppercase">Edit Activity Record</h3>
                </div>
                <button onClick={() => setActivityEditModal(null)} className="text-slate-300 hover:text-white cursor-pointer">
                  <X className="w-4 h-4" />
                </button>
              </div>
              <form onSubmit={handleSaveActivityEdit} className="p-5 flex flex-col gap-4 text-xs font-semibold">
                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Student</label>
                  <select
                    value={activityEditModal.student_id}
                    onChange={(e) => setActivityEditModal({ ...activityEditModal, student_id: e.target.value })}
                    className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                    required
                  >
                    {students.map(s => (
                      <option key={s.id} value={s.id}>{s.name} ({s.class.toUpperCase()})</option>
                    ))}
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Activity Type</label>
                    <input
                      type="text"
                      value={activityEditModal.event_type}
                      onChange={(e) => setActivityEditModal({ ...activityEditModal, event_type: e.target.value })}
                      placeholder="e.g. tally, star, spot fine"
                      className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                      required
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Amount</label>
                    <input
                      type="number"
                      value={activityEditModal.amount}
                      onChange={(e) => setActivityEditModal({ ...activityEditModal, amount: Number(e.target.value) })}
                      className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                      required
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Date</label>
                  <input
                    type="date"
                    value={activityEditModal.date}
                    onChange={(e) => setActivityEditModal({ ...activityEditModal, date: e.target.value })}
                    className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                    required
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Reason / Note</label>
                  <input
                    type="text"
                    value={activityEditModal.reason || ''}
                    onChange={(e) => setActivityEditModal({ ...activityEditModal, reason: e.target.value })}
                    placeholder="Enter reason..."
                    className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                  />
                </div>

                <div className="flex gap-2 justify-end pt-2 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setActivityEditModal(null)}
                    className="px-4 py-2 bg-slate-100 text-slate-600 font-bold rounded-xl cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-4 py-2 bg-[#1A365D] hover:bg-blue-900 text-white font-bold rounded-xl shadow-sm cursor-pointer"
                  >
                    Save Changes
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* Add New Activity Modal */}
        {showAddActivityModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[200] animate-fade-in">
            <div className="bg-white rounded-2xl w-full max-w-md shadow-2xl border border-slate-200 overflow-hidden flex flex-col scale-in">
              <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-emerald-700 text-white">
                <div className="flex items-center gap-2">
                  <PlusCircle className="w-4 h-4" />
                  <h3 className="font-extrabold tracking-wider text-sm uppercase">Add Activity Record</h3>
                </div>
                <button onClick={() => setShowAddActivityModal(false)} className="text-emerald-100 hover:text-white cursor-pointer">
                  <X className="w-4 h-4" />
                </button>
              </div>
              <form onSubmit={handleCreateActivity} className="p-5 flex flex-col gap-4 text-xs font-semibold">
                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Select Student</label>
                  <select
                    value={newActivityForm.student_id}
                    onChange={(e) => setNewActivityForm({ ...newActivityForm, student_id: e.target.value })}
                    className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                    required
                  >
                    <option value="" disabled>Choose student...</option>
                    {students.map(s => (
                      <option key={s.id} value={s.id}>{s.name} ({s.class.toUpperCase()})</option>
                    ))}
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Activity Type</label>
                    <select
                      value={newActivityForm.event_type}
                      onChange={(e) => setNewActivityForm({ ...newActivityForm, event_type: e.target.value })}
                      className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                      required
                    >
                      <option value="tally">Tally</option>
                      <option value="star">Star</option>
                      <option value="spot fine">Spot Fine</option>
                      <option value="n&o">N&O Tally</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Amount</label>
                    <input
                      type="number"
                      value={newActivityForm.amount}
                      onChange={(e) => setNewActivityForm({ ...newActivityForm, amount: Number(e.target.value) })}
                      className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                      required
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Date</label>
                  <input
                    type="date"
                    value={newActivityForm.date}
                    onChange={(e) => setNewActivityForm({ ...newActivityForm, date: e.target.value })}
                    className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                    required
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">Reason / Note</label>
                  <input
                    type="text"
                    value={newActivityForm.reason}
                    onChange={(e) => setNewActivityForm({ ...newActivityForm, reason: e.target.value })}
                    placeholder="Enter reason..."
                    className="w-full p-3 border border-slate-200 rounded-xl font-bold text-slate-800 bg-slate-50"
                  />
                </div>

                <div className="flex gap-2 justify-end pt-2 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setShowAddActivityModal(false)}
                    className="px-4 py-2 bg-slate-100 text-slate-600 font-bold rounded-xl cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold rounded-xl shadow-sm cursor-pointer"
                  >
                    Add Record
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* ISSUE PHONE PASS MODAL */}
        {showIssuePhonePassModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[210] animate-fade-in font-sans">
            <div className="bg-white rounded-3xl w-full max-w-lg max-h-[90vh] shadow-2xl border border-slate-100 overflow-hidden flex flex-col p-6 gap-4" onClick={e => e.stopPropagation()}>

              {/* Modal Header */}
              <div className="flex items-center justify-between border-b border-slate-100 pb-3 shrink-0">
                <div className="flex items-center gap-2 text-emerald-700">
                  <Smartphone className="w-6 h-6" />
                  <h3 className="font-black text-lg uppercase">Issue School Phone Pass</h3>
                </div>
                <button onClick={() => setShowIssuePhonePassModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
              </div>

              <div className="overflow-y-auto pr-1 flex flex-col gap-4">

              {phonePassStep === 1 ? (
                /* STEP 1: SELECT STUDENT */
                <div className="flex flex-col gap-4">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-extrabold text-slate-500 uppercase tracking-wider">
                      Step 1: Select School Student
                    </span>
                    <span className="text-[11px] font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md">
                      Showing School Category Students
                    </span>
                  </div>

                  <div className="relative">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={phonePassStudentSearch}
                      onChange={e => setPhonePassStudentSearch(e.target.value)}
                      placeholder="Search student by name or class..."
                      className="w-full py-2.5 pl-10 pr-4 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-500"
                      autoFocus
                    />
                  </div>

                  <div className="max-h-64 overflow-y-auto divide-y divide-slate-100 border border-slate-200 rounded-2xl bg-white">
                    {students
                      .filter(s => (s.phoneType || 'school') === 'school')
                      .filter(s => !phonePassStudentSearch ||
                        s.name.toLowerCase().includes(phonePassStudentSearch.toLowerCase()) ||
                        s.class.toLowerCase().includes(phonePassStudentSearch.toLowerCase()) ||
                        (s.registerNumber && s.registerNumber.includes(phonePassStudentSearch))
                      )
                      .map(st => {
                        const hasActivePass = phonePasses.some(p => String(p.studentId) === String(st.id) && p.status === 'OUT');

                        return (
                          <button
                            key={st.id}
                            type="button"
                            onClick={() => {
                              setPhonePassSelectedStudent(st);
                              setPhonePassStep(2);
                            }}
                            className="w-full text-left p-3 hover:bg-emerald-50 flex items-center justify-between transition-colors group"
                          >
                            <div className="flex items-center gap-3">
                              <div className="w-8 h-8 rounded-xl bg-emerald-100 text-emerald-800 font-extrabold text-xs flex items-center justify-center">
                                {st.name.charAt(0)}
                              </div>
                              <div>
                                <div className="flex items-center gap-2">
                                  <h4 className="font-extrabold text-xs text-slate-800 group-hover:text-emerald-800">{st.name}</h4>
                                  {st.registerNumber && (
                                    <span className="font-mono text-[10px] font-black text-indigo-700 bg-indigo-50 px-1.5 py-0.5 rounded border border-indigo-100">
                                      #{st.registerNumber}
                                    </span>
                                  )}
                                </div>
                                <span className="text-[10px] font-bold text-slate-400 uppercase">Class {st.class}</span>
                              </div>
                            </div>

                            <div>
                              {hasActivePass ? (
                                <span className="text-[10px] font-black bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full uppercase">
                                  Active OUT Pass
                                </span>
                              ) : (
                                <span className="text-xs font-extrabold text-emerald-700 group-hover:translate-x-1 transition-transform block">
                                  Select →
                                </span>
                              )}
                            </div>
                          </button>
                        );
                      })}
                  </div>

                  <div className="pt-2">
                    <button
                      type="button"
                      onClick={() => setShowIssuePhonePassModal(false)}
                      className="w-full py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider"
                    >
                      Cancel
                    </button>
                  </div>
                </div>
              ) : (
                /* STEP 2: ISSUE FORM WITH PRE-FILLED READ-ONLY DETAILS & DURATION */
                <form onSubmit={handleIssuePhonePassSubmit} className="flex flex-col gap-4">

                  {/* Phone Model missing warning check */}
                  {(!phonePassSelectedStudent?.phoneModel || !phonePassSelectedStudent?.phoneModel.trim()) && (
                    <div className="p-3 bg-amber-100 border border-amber-300 rounded-2xl text-amber-900 text-xs font-extrabold flex items-center gap-2">
                      <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0" />
                      <span>⚠️ Warning: Phone Name / Model is missing for this student. You must register a Phone Name before issuing a pass.</span>
                    </div>
                  )}

                  {/* Active Pass Warning check */}
                  {phonePasses.some(p => String(p.studentId) === String(phonePassSelectedStudent?.id) && p.status === 'OUT') && (
                    <div className="p-3 bg-rose-100 border border-rose-300 rounded-2xl text-rose-800 text-xs font-extrabold flex items-center gap-2">
                      <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0" />
                      <span>⚠️ Warning: Student already has an active OUT pass! Previous pass must be marked IN before issuing a new pass.</span>
                    </div>
                  )}

                  {/* Pre-filled Read-only Student Details */}
                  <div className="bg-slate-50 p-3.5 rounded-2xl border border-slate-200 flex flex-col gap-2">
                    <div className="flex items-center justify-between">
                      <span className="text-[10px] font-extrabold text-slate-400 uppercase">Student Name (Read Only)</span>
                      <button
                        type="button"
                        onClick={() => setPhonePassStep(1)}
                        className="text-[10px] font-extrabold text-emerald-700 hover:underline"
                      >
                        Change Student
                      </button>
                    </div>
                    <div className="text-sm font-black text-slate-800 flex items-center gap-2">
                      <span>{phonePassSelectedStudent?.name}</span>
                      {phonePassSelectedStudent?.registerNumber && (
                        <span className="font-mono text-xs font-black text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-100">
                          #{phonePassSelectedStudent?.registerNumber}
                        </span>
                      )}
                    </div>
                    <div className="flex items-center gap-2 mt-0.5">
                      <span className="text-[10px] font-extrabold bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-md uppercase">
                        Class {phonePassSelectedStudent?.class}
                      </span>
                      <span className="text-[10px] font-extrabold bg-slate-200 text-slate-700 px-2 py-0.5 rounded-md uppercase">
                        School Student
                      </span>
                    </div>
                  </div>

                  {/* Phone Model / Device Details Input */}
                  <div>
                    <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1">
                      Phone Name / Model
                    </label>
                    <input
                      type="text"
                      value={phonePassSelectedStudent?.phoneModel || ''}
                      onChange={e => {
                        const val = e.target.value;
                        setPhonePassSelectedStudent(prev => prev ? { ...prev, phoneModel: val } : null);
                        if (phonePassSelectedStudent?.id) {
                          handleUpdateStudentPhoneDetails(
                            phonePassSelectedStudent.id,
                            phonePassSelectedStudent.phoneType || 'school',
                            val,
                            phonePassSelectedStudent.registerNumber || ''
                          );
                        }
                      }}
                      placeholder="Enter phone name/model (e.g. Samsung A52, iPhone 13)..."
                      className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  {/* Starting Time (Auto Real-Time) */}
                  <div>
                    <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1">
                      Pass Start Time (Real-Time Auto Set)
                    </label>
                    <input
                      type="text"
                      value={`Current Time: ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true })}`}
                      disabled
                      className="w-full p-3 bg-slate-100 border border-slate-200 rounded-xl font-extrabold text-xs text-slate-700 cursor-not-allowed"
                    />
                  </div>

                  {/* Select Return Date & Return Time */}
                  <div className="flex flex-col gap-3 bg-emerald-50/60 p-3.5 rounded-2xl border border-emerald-200">
                    <div className="flex items-center justify-between">
                      <label className="text-xs font-black text-emerald-900 uppercase tracking-wider flex items-center gap-1.5">
                        <Calendar className="w-4 h-4 text-emerald-600" />
                        Set Return Date & Time
                      </label>
                    </div>

                    {/* Return Date Input */}
                    <div>
                      <span className="text-[11px] font-extrabold text-slate-600 block mb-1">Return Date:</span>
                      <input
                        type="date"
                        value={phonePassReturnDate}
                        onChange={e => setPhonePassReturnDate(e.target.value)}
                        className="w-full p-2.5 bg-white border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-500 shadow-xs"
                      />
                    </div>

                    {/* Allowed Duration / Time Options */}
                    <div>
                      <span className="text-[11px] font-extrabold text-slate-600 block mb-1.5">Return Time Presets:</span>
                      <div className="grid grid-cols-4 gap-1.5 mb-2">
                        {[
                          { mins: 30, label: '30 Mins' },
                          { mins: 60, label: '1 Hour' },
                          { mins: 120, label: '2 Hours' },
                          { mins: 180, label: '3 Hours' }
                        ].map(item => (
                          <button
                            key={item.mins}
                            type="button"
                            onClick={() => {
                              setPhonePassAllowedMins(item.mins);
                              setPhonePassCustomAllowedTime('');
                            }}
                            className={`py-2 px-2 rounded-xl text-xs font-extrabold transition-all border ${phonePassAllowedMins === item.mins && !phonePassCustomAllowedTime
                                ? 'bg-emerald-600 text-white border-emerald-600 shadow-xs'
                                : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                              }`}
                          >
                            {item.label}
                          </button>
                        ))}
                      </div>

                      <div className="flex items-center justify-between gap-2 bg-white p-2 rounded-xl border border-slate-200">
                        <span className="text-xs font-extrabold text-slate-600">Custom Return Time:</span>
                        <input
                          type="time"
                          value={phonePassCustomAllowedTime}
                          onChange={e => setPhonePassCustomAllowedTime(e.target.value)}
                          className="py-1 px-3 bg-slate-50 border border-slate-200 rounded-lg font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-500"
                        />
                      </div>
                    </div>
                  </div>

                  {/* Reason for Permitting Pass */}
                  <div>
                    <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1.5">
                      Reason for Permitting Pass
                    </label>
                    <input
                      type="text"
                      value={phonePassReason}
                      onChange={e => setPhonePassReason(e.target.value)}
                      placeholder="Enter reason (e.g. Calling Family)..."
                      className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-emerald-500"
                      required
                    />
                    <div className="flex flex-wrap gap-1 mt-2">
                      {['Calling Home', 'Study / Project', 'Emergency Call', 'Doctor Appt', 'Personal'].map(r => (
                        <button
                          key={r}
                          type="button"
                          onClick={() => setPhonePassReason(r)}
                          className="text-[10px] font-extrabold bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-md transition-colors"
                        >
                          {r}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Footer Submit Buttons */}
                  <div className="flex items-center gap-3 pt-2">
                    <button
                      type="button"
                      onClick={() => setPhonePassStep(1)}
                      className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider"
                    >
                      Back
                    </button>
                    <button
                      type="submit"
                      disabled={!phonePassSelectedStudent?.phoneModel || !phonePassSelectedStudent?.phoneModel.trim() || phonePasses.some(p => String(p.studentId) === String(phonePassSelectedStudent?.id) && p.status === 'OUT')}
                      className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 disabled:cursor-not-allowed text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-1.5"
                    >
                      <Smartphone className="w-4 h-4" />
                      <span>Issue Pass</span>
                    </button>
                  </div>
                </form>
              )}

              </div>
            </div>
          </div>
        )}

        {/* --- MONTHLY LEAVE BULK PHONE PASS MODAL --- */}
        {showMonthlyLeaveModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-fade-in">
            <div className="bg-white w-full max-w-lg max-h-[90vh] rounded-3xl p-5 shadow-2xl border border-slate-100 flex flex-col gap-4 animate-scale-up overflow-y-auto" onClick={e => e.stopPropagation()}>
              {/* Modal Header */}
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2 text-indigo-700">
                  <Calendar className="w-6 h-6" />
                  <h3 className="font-black text-lg uppercase">Issue Monthly Leave Phone Passes</h3>
                </div>
                <button onClick={() => setShowMonthlyLeaveModal(false)} className="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
              </div>

              <form onSubmit={handleIssueMonthlyLeavePasses} className="flex flex-col gap-4">
                {/* Info Alert */}
                <div className="p-3 bg-indigo-50 border border-indigo-100 rounded-2xl text-indigo-900 text-xs font-semibold flex flex-col gap-1">
                  <div className="font-extrabold flex items-center gap-1.5 text-indigo-800">
                    <Sparkles className="w-4 h-4 text-indigo-600 shrink-0" />
                    <span>Automatic Bulk Issuance</span>
                  </div>
                  <p>This will issue Phone Passes to <strong>ALL students</strong> except those who are currently marked as <span className="text-rose-600 font-extrabold">Ineligible</span> or already have an active pass.</p>
                </div>

                {/* Class Filter Selection */}
                <div>
                  <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1">
                    Select Target Class
                  </label>
                  <select
                    value={monthlyLeaveClassFilter}
                    onChange={e => setMonthlyLeaveClassFilter(e.target.value)}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 uppercase focus:outline-none focus:border-indigo-500"
                  >
                    <option value="all">All Classes ({students.filter(s => !s.ineligible).length} Eligible Students)</option>
                    {CLASSES.map(c => {
                      const count = students.filter(s => !s.ineligible && s.class.toLowerCase() === c.toLowerCase()).length;
                      return (
                        <option key={c} value={c}>Class {c.toUpperCase()} ({count} Eligible Students)</option>
                      );
                    })}
                  </select>
                </div>

                {/* Expected Return Date & Time Selection */}
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1">
                      Return Date
                    </label>
                    <input
                      type="date"
                      value={monthlyLeaveDate}
                      onChange={e => setMonthlyLeaveDate(e.target.value)}
                      className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-indigo-500"
                      required
                    />
                  </div>

                  <div>
                    <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1">
                      Return Time
                    </label>
                    <input
                      type="time"
                      value={monthlyLeaveTime}
                      onChange={e => setMonthlyLeaveTime(e.target.value)}
                      className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-indigo-500"
                      required
                    />
                  </div>
                </div>

                {/* Reason for Monthly Leave */}
                <div>
                  <label className="text-xs font-extrabold text-slate-500 uppercase tracking-wider block mb-1">
                    Pass Reason / Description
                  </label>
                  <input
                    type="text"
                    value={monthlyLeaveReason}
                    onChange={e => setMonthlyLeaveReason(e.target.value)}
                    placeholder="e.g. Monthly Leave August 2026"
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:outline-none focus:border-indigo-500"
                    required
                  />
                </div>

                {/* Summary of Eligible vs Ineligible */}
                {(() => {
                  const totalInClass = students.filter(s => monthlyLeaveClassFilter === 'all' || s.class.toLowerCase() === monthlyLeaveClassFilter.toLowerCase());
                  const ineligibleCount = totalInClass.filter(s => s.ineligible).length;
                  const eligibleCount = totalInClass.filter(s => !s.ineligible).length;

                  return (
                    <div className="grid grid-cols-2 gap-2 text-center text-xs font-extrabold p-2.5 bg-slate-50 rounded-2xl border border-slate-200">
                      <div className="p-2 bg-emerald-50 text-emerald-800 rounded-xl border border-emerald-100">
                        <div className="text-base font-black">{eligibleCount}</div>
                        <div className="text-[10px] uppercase">Will Receive Pass</div>
                      </div>
                      <div className="p-2 bg-rose-50 text-rose-800 rounded-xl border border-rose-100">
                        <div className="text-base font-black">{ineligibleCount}</div>
                        <div className="text-[10px] uppercase">Ineligible (Skipped)</div>
                      </div>
                    </div>
                  );
                })()}

                {/* Action Buttons */}
                <div className="flex items-center gap-3 pt-2">
                  <button
                    type="button"
                    onClick={() => setShowMonthlyLeaveModal(false)}
                    className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-extrabold text-xs uppercase tracking-wider"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="flex-1 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-extrabold text-xs uppercase tracking-wider shadow-md transition-all active:scale-95 flex items-center justify-center gap-1.5"
                  >
                    <Calendar className="w-4 h-4" />
                    <span>Issue All Passes</span>
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}

