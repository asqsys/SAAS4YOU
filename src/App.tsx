/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import QRCode from 'qrcode';
import { 
  FileSpreadsheet, 
  Upload, 
  FileText, 
  CheckCircle2, 
  AlertCircle, 
  Download,
  Building2,
  Package,
  Calculator,
  Archive,
  Users,
  Check,
  HelpCircle,
  X
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

const VERSION = "1.2.0";

// --- Types ---

interface Company {
  id: string;
  name: string;
  address: string;
  email: string;
  vat: string;
  ice: string;
  bankAccount1: string;
  bankAccount2: string;
}

interface ThirdParty {
  id: string;
  name: string;
  address?: string;
  email?: string;
  vatNumber?: string;
  ice?: string;
}

interface Service {
  id: string;
  ref: string;
  description: string;
  matchKey: string; // prenom + premiere lettre nom
  unitPrice?: number; // Added for the check
}

interface BillingLine {
  clientName: string;
  serviceUser: string; // The person name to match with service
  date: string;
  quantity: number;
  tjmAO: number;
  tjmPortage: number;
  amount: number; // Total HT from file
}

interface InvoiceData {
  client: ThirdParty;
  lines: {
    service: Service;
    quantity: number;
    unitPrice: number;
    total: number;
    date: string;
  }[];
  totalHT: number;
  totalTTC: number;
}

// --- Main Component ---

export default function App() {
  const [thirdPartyData, setThirdPartyData] = useState<ThirdParty[]>([]);
  const [servicesData, setServicesData] = useState<Service[]>([]);
  const [billingData, setBillingData] = useState<BillingLine[]>([]);
  
  const [filesUploaded, setFilesUploaded] = useState({
    thirdParty: false,
    services: false,
    billing: false,
    company: false
  });

  const [companies, setCompanies] = useState<Company[]>(() => {
    const saved = localStorage.getItem('asqsys_companies');
    return saved ? JSON.parse(saved) : [];
  });

  const [selectedCompanyId, setSelectedCompanyId] = useState<string | null>(() => {
    return localStorage.getItem('asqsys_selected_company_id');
  });

  // Derived selected company
  const companyInfo = companies.find(c => c.id === selectedCompanyId) || {
    id: "",
    name: "",
    address: "",
    email: "",
    vat: "",
    ice: "",
    bankAccount1: "",
    bankAccount2: ""
  };

  useEffect(() => {
    localStorage.setItem('asqsys_companies', JSON.stringify(companies));
    // If no company selected but list not empty, select first
    if (!selectedCompanyId && companies.length > 0) {
      setSelectedCompanyId(companies[0].id);
    }
  }, [companies, selectedCompanyId]);

  useEffect(() => {
    if (selectedCompanyId) {
      localStorage.setItem('asqsys_selected_company_id', selectedCompanyId);
      setFilesUploaded(prev => ({ ...prev, company: true }));

      // Ensure at least one valid bank account is selected for the new company
      const company = companies.find(c => c.id === selectedCompanyId);
      if (company) {
        const hasAcc1 = !!company.bankAccount1;
        const hasAcc2 = !!company.bankAccount2;
        const isAcc1Valid = selectedBankAccounts.account1 && hasAcc1;
        const isAcc2Valid = selectedBankAccounts.account2 && hasAcc2;
        
        if (!isAcc1Valid && !isAcc2Valid && (hasAcc1 || hasAcc2)) {
          setSelectedBankAccounts({
            account1: hasAcc1,
            account2: !hasAcc1 && hasAcc2
          });
        }
      }
    } else {
      localStorage.removeItem('asqsys_selected_company_id');
      setFilesUploaded(prev => ({ ...prev, company: false }));
    }
  }, [selectedCompanyId, companies]);

  const [selectedBankAccounts, setSelectedBankAccounts] = useState({
    account1: true,
    account2: false
  });

  const [showSettings, setShowSettings] = useState(false);
  const [showFAQ, setShowFAQ] = useState(false);

  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [creditDuration, setCreditDuration] = useState(30);
  const [detectedClients, setDetectedClients] = useState<string[]>([]);
  const [selectedClients, setSelectedClients] = useState<Set<string>>(new Set());
  const [billingOptions, setBillingOptions] = useState<Set<'GLOBAL' | 'RATE_A' | 'RATE_B'>>(new Set(['GLOBAL']));
  const [invoiceNote, setInvoiceNote] = useState("");
  const [invoicePeriod, setInvoicePeriod] = useState("");

  const getLastDaysOfMonths = () => {
    const dates = [];
    const now = new Date();
    for (let i = -12; i <= 6; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() + i + 1, 0);
      dates.push(d.toISOString().split('T')[0]);
    }
    return dates.sort().reverse();
  };

  const [selectedInvoiceDate, setSelectedInvoiceDate] = useState(getLastDaysOfMonths()[6] || new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0).toISOString().split('T')[0]);
  const [invoiceReference, setInvoiceReference] = useState(`FAC-${Date.now().toString().slice(-6)}`);

  useEffect(() => {
    if (selectedInvoiceDate) {
      const date = new Date(selectedInvoiceDate);
      const period = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      setInvoicePeriod(period);
    }
  }, [selectedInvoiceDate]);

  const formatDH = (num: number) => {
    const parts = num.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).split(',');
    const integerPart = parts[0].replace(/\s/g, ' ');
    return `${integerPart},${parts[1]} DH`;
  };

  const getMachineFingerprint = () => {
    try {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) return 'NO-CANVAS';
      
      const txt = 'ASQSYS-VERIFY-1234567890-!@#$%^&*()';
      ctx.textBaseline = "top";
      ctx.font = "14px 'Arial'";
      ctx.textBaseline = "alphabetic";
      ctx.fillStyle = "#f60";
      ctx.fillRect(125,1,62,20);
      ctx.fillStyle = "#069";
      ctx.fillText(txt, 2, 15);
      ctx.fillStyle = "rgba(102, 204, 0, 0.7)";
      ctx.fillText(txt, 4, 17);
      
      return canvas.toDataURL().slice(-50); 
    } catch (e) {
      return 'FINGERPRINT-ERROR';
    }
  };

  const getPersistentId = () => {
    let id = localStorage.getItem('asqsys_machine_id');
    if (!id) {
      if (typeof crypto !== 'undefined' && crypto.randomUUID) {
        id = crypto.randomUUID();
      } else {
        // Fallback for non-secure contexts (HTTP)
        id = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
          const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
          return v.toString(16);
        });
      }
      localStorage.setItem('asqsys_machine_id', id);
    }
    return id;
  };

  const incrementReference = (ref: string) => {
    const match = ref.match(/^(.*?)(\d+)$/);
    if (!match) return ref + "-1";
    const prefix = match[1];
    const numberStr = match[2];
    const nextNumber = parseInt(numberStr) + 1;
    const nextNumberStr = nextNumber.toString().padStart(numberStr.length, '0');
    return prefix + nextNumberStr;
  };

  const numberToWordsFR = (num: number) => {
    const units = ['', 'un', 'deux', 'trois', 'quatre', 'cinq', 'six', 'sept', 'huit', 'neuf'];
    const teens = ['dix', 'onze', 'douze', 'treize', 'quatorze', 'quinze', 'seize', 'dix-sept', 'dix-huit', 'dix-neuf'];
    const tens = ['', 'dix', 'vingt', 'trente', 'quarante', 'cinquante', 'soixante', 'soixante-dix', 'quatre-vingt', 'quatre-vingt-dix'];

    const convert = (n: number): string => {
      if (n === 0) return '';
      if (n < 10) return units[n];
      if (n < 20) return teens[n - 10];
      if (n < 100) {
        const t = Math.floor(n / 10);
        const u = n % 10;
        if (t === 7 || t === 9) {
          return tens[t - 1] + (u === 1 ? ' et ' : '-') + convert(u + 10);
        }
        return tens[t] + (u === 1 ? ' et ' : (u > 0 ? '-' : '')) + units[u];
      }
      if (n < 1000) {
        const c = Math.floor(n / 100);
        const r = n % 100;
        if (c === 1) return 'cent' + (r > 0 ? ' ' + convert(r) : '');
        return units[c] + ' cent' + (r === 0 ? 's' : ' ' + convert(r));
      }
      if (n < 1000000) {
        const m = Math.floor(n / 1000);
        const r = n % 1000;
        const mille = m === 1 ? 'mille' : convert(m) + ' mille';
        return mille + (r > 0 ? ' ' + convert(r) : '');
      }
      if (n < 1000000000) {
        const mill = Math.floor(n / 1000000);
        const r = n % 1000000;
        const million = mill === 1 ? 'un million' : convert(mill) + ' millions';
        return million + (r > 0 ? ' ' + convert(r) : '');
      }
      return n.toString();
    };

    const integerPart = Math.floor(num);
    const decimalPart = Math.round((num - integerPart) * 100);

    if (integerPart === 0 && decimalPart === 0) return 'Zéro dirham';

    let result = '';
    if (integerPart > 0) {
      const needsDe = integerPart >= 1000000 && integerPart % 1000000 === 0;
      result = convert(integerPart) + (needsDe ? ' de' : '') + (integerPart > 1 ? ' dirhams' : ' dirham');
    }
    
    if (decimalPart > 0) {
      const cents = convert(decimalPart) + (decimalPart > 1 ? ' centimes' : ' centime');
      result = result ? result + ' et ' + cents : cents;
    }
    return result.charAt(0).toUpperCase() + result.slice(1);
  };

  // --- File Parsing ---

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>, type: 'thirdParty' | 'services' | 'billing' | 'company') => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];

      const findValue = (row: any, keys: string[]) => {
        const rowKeys = Object.keys(row);
        for (const key of keys) {
          if (row[key] !== undefined) return row[key];
          const foundKey = rowKeys.find(rk => rk.toLowerCase().trim() === key.toLowerCase().trim());
          if (foundKey) return row[foundKey];
        }
        return undefined;
      };

      if (type === 'company') {
        if (jsonData.length === 0) {
          setError("Le fichier société est vide.");
          return;
        }
        const row = jsonData[0];
        const name = String(findValue(row, ['Name', 'Nom', 'Raison Sociale']) || "VOTRE ENTREPRISE");
        const addr = String(findValue(row, ['Address', 'Adresse']) || "");
        const cp = String(findValue(row, ['Postcode', 'Code Postal', 'CP']) || "");
        const city = String(findValue(row, ['City', 'Ville']) || "");
        const country = String(findValue(row, ['Country', 'Pays']) || "");
        
        const bank1 = String(findValue(row, ['Payment bank account 1', 'Bank 1', 'Compte 1']) || "");
        const bank2 = String(findValue(row, ['Payment bank account 2', 'Bank 2', 'Compte 2']) || "");
        
        const ice = String(findValue(row, ['ICE', 'Identifiant Commun Entreprise']) || "");
        
        const newCompany: Company = {
          id: Date.now().toString(),
          name,
          address: `${addr}${cp ? ', ' + cp : ''}${city ? ' ' + city : ''}${country ? ', ' + country : ''}`.trim() || "Adresse non renseignée",
          email: String(findValue(row, ['Email', 'Courriel']) || "contact@entreprise.com"),
          vat: String(findValue(row, ['VAT ID', 'TVA', 'Numéro TVA', 'VAT']) || "Non renseigné"),
          ice,
          bankAccount1: bank1,
          bankAccount2: bank2
        };

        setCompanies(prev => {
          const existingIndex = prev.findIndex(c => c.name.toLowerCase() === name.toLowerCase());
          let updatedCompanies;
          let targetId;
          
          if (existingIndex >= 0) {
            targetId = prev[existingIndex].id;
            updatedCompanies = [...prev];
            updatedCompanies[existingIndex] = { ...newCompany, id: targetId };
          } else {
            targetId = newCompany.id;
            updatedCompanies = [...prev, newCompany];
          }
          
          setSelectedCompanyId(targetId);
          return updatedCompanies;
        });

        setFilesUploaded(prev => ({ ...prev, company: true }));
        setError(null);
      } else if (type === 'thirdParty') {
        const parsed = jsonData.map(row => {
          const addr = String(findValue(row, ['Address', 'Adresse']) || "");
          const cp = String(findValue(row, ['Postcode', 'Code Postal', 'CP']) || "");
          const city = String(findValue(row, ['City', 'Ville']) || "");
          const country = String(findValue(row, ['Country', 'Pays']) || "");

          return {
            id: String(findValue(row, ['Customer Code', 'ID', 'Reference', 'Ref', 'Code']) || ''),
            name: String(findValue(row, ['Name', 'Nom', 'Client', 'Entreprise', 'Société']) || ''),
            address: `${addr}${cp ? ', ' + cp : ''}${city ? ' ' + city : ''}${country ? ', ' + country : ''}`.trim(),
            email: String(findValue(row, ['Email', 'Courriel']) || ''),
            vatNumber: String(findValue(row, ['VAT ID', 'VAT', 'TVA', 'N° TVA']) || ''),
            ice: String(findValue(row, ['ICE', 'Identifiant Commun Entreprise']) || '')
          };
        });

        if (parsed.length === 0 || !parsed.some(p => p.name)) {
          setError("Le fichier tiers semble invalide ou ne contient pas de colonne 'Name' ou 'Nom'.");
          return;
        }

        setThirdPartyData(parsed);
        setFilesUploaded(prev => ({ ...prev, thirdParty: true }));
        setError(null);
      } else if (type === 'services') {
        const parseNum = (val: any) => {
          if (typeof val === 'number') return val;
          if (typeof val === 'string') {
            const cleaned = val.replace(',', '.').replace(/[^-0-9.]/g, '');
            return cleaned ? Number(cleaned) : 0;
          }
          return 0;
        };

        const parsed = jsonData.map(row => {
          const id = String(findValue(row, ['ID', 'Ref.', 'Reference', 'Ref', 'Code']) || '');
          const ref = String(findValue(row, ['Ref.', 'Reference', 'Ref', 'Code', 'ID']) || '');
          const description = String(findValue(row, ['Label', 'Description', 'Service', 'Prestation', 'Libellé']) || '');
          const unitPrice = parseNum(findValue(row, ['P.U. (excl.)', 'Unit price (excl.)', 'Prix Unitaire', 'PU', 'Taux', 'TJM', 'Prix', 'Tarif']));
          return {
            id,
            ref,
            description,
            matchKey: ref.toUpperCase(),
            unitPrice: unitPrice || undefined
          };
        });

        if (parsed.length === 0 || (!parsed.some(p => p.id) && !parsed.some(p => p.ref))) {
          setError("Le fichier de services semble invalide ou ne contient pas de colonne 'ID' ou 'Ref.'.");
          return;
        }

        setServicesData(parsed);
        setFilesUploaded(prev => ({ ...prev, services: true }));
        setError(null);
      } else if (type === 'billing') {
        const parseNum = (val: any) => {
          if (typeof val === 'number') return val;
          if (typeof val === 'string') {
            const cleaned = val.replace(',', '.').replace(/[^-0-9.]/g, '');
            return cleaned ? Number(cleaned) : 0;
          }
          return 0;
        };

        const parsed = jsonData.map(row => {
          // Specific columns provided by user
          const tjmAO = parseNum(findValue(row, ['TJM A.O', 'TJM AO', 'AO', 'TJM A.O HT', 'TJM AO HT', 'Tarif A']));
          const tjmPortage = parseNum(findValue(row, ['TJM Portage', 'Portage', 'TJM Portage HT', 'Tarif B']));
          const quantity = parseNum(findValue(row, ['Nbre Jours', 'Nbre de jours', 'Quantite', 'Nb', 'Heures', 'Jours', 'Qté', 'Quantity', 'Qty', 'Units']));
          const amount = parseNum(findValue(row, ['Total HT', 'Factures HT', 'Montant', 'Total', 'HT', 'Net', 'Amount', 'Total HT Net', 'Total Net HT']));
          
          return {
            clientName: String(findValue(row, ['Client', 'Entreprise', 'Société', 'Tiers', 'Customer', 'Third Party', 'Thirdparty']) || ''),
            serviceUser: String(findValue(row, ['Description', 'Consultant', 'Collaborateur', 'Nom', 'Prenom Nom', 'Salarié', 'User', 'Resource', 'Staff']) || ''),
            amount: amount,
            date: String(findValue(row, ['Date', 'Commentaires', 'Période', 'Mois', 'Month', 'Period']) || ''),
            quantity: quantity,
            tjmAO,
            tjmPortage
          };
        }).filter(row => row.serviceUser && (row.quantity > 0 || row.amount > 0) && (row.tjmAO > 0 || row.tjmPortage > 0 || row.amount > 0));
        
        if (jsonData.length === 0) {
          setError(`Le fichier ${type} semble être vide.`);
          return;
        }

        if (parsed.length === 0) {
          const sampleRow = jsonData[0] || {};
          const keys = Object.keys(sampleRow).join(', ');
          setError(`Aucune donnée valide trouvée dans le fichier de facturation.\n\nColonnes détectées : ${keys}\n\nVérifiez que les colonnes 'TJM A.O', 'TJM Portage' ou 'Total HT' sont bien remplies et que la colonne 'Description' (ou 'Consultant') contient des noms.`);
          return;
        }

        setBillingData(parsed);
        const clients = Array.from(new Set(parsed.map(p => {
          let name = p.clientName.trim();
          if (!name && thirdPartyData.length > 0) name = thirdPartyData[0].name;
          return name || "Client Inconnu";
        })));
        setDetectedClients(clients);
        setSelectedClients(new Set(clients));
        setFilesUploaded(prev => ({ ...prev, billing: true }));
        setError(null);
      }
    } catch (err) {
      console.error(err);
      setError(`Erreur lors de la lecture du fichier ${type}`);
    }
  };

  // --- Logic ---

  const generateInvoices = async () => {
    if (billingData.length === 0) {
      setError("Aucune donnée de facturation trouvée.");
      return;
    }

    if (filesUploaded.company && (companyInfo.bankAccount1 || companyInfo.bankAccount2)) {
      if (!selectedBankAccounts.account1 && !selectedBankAccounts.account2) {
        setError("Veuillez sélectionner au moins un compte bancaire pour le paiement.");
        return;
      }
    }

    setProcessing(true);
    try {
      // 1. Validation check before anything else
      const validationErrors: string[] = [];
      
      billingData.forEach(line => {
        const nameParts = line.serviceUser.trim().split(/\s+/);
        if (nameParts.length < 2) return;

        const p1 = nameParts[0].toUpperCase();
        const p2 = nameParts[nameParts.length - 1].toUpperCase();
        const firstLetter = p2.charAt(0);
        
        // Match keys: Prename-FirstLetterOfName
        const matchKey = `${p1}-${firstLetter}`;

        // Check Rate A (AO)
        if ((billingOptions.has('GLOBAL') || billingOptions.has('RATE_A')) && line.tjmAO > 0) {
          const serviceAO = servicesData.find(s => s.ref.toUpperCase().includes('-AA-') && s.ref.toUpperCase().includes(matchKey));
          if (serviceAO && serviceAO.unitPrice !== undefined && Math.abs(serviceAO.unitPrice - line.tjmAO) > 0.01) {
            validationErrors.push(`Écart Tarif A (AO) pour ${line.serviceUser}: Facturation (${line.tjmAO} DH) vs Services (${serviceAO.unitPrice} DH)`);
          }
        }

        // Check Rate B (Portage)
        if ((billingOptions.has('GLOBAL') || billingOptions.has('RATE_B')) && line.tjmPortage > 0) {
          const servicePortage = servicesData.find(s => s.ref.toUpperCase().includes('-PR-') && s.ref.toUpperCase().includes(matchKey));
          if (servicePortage && servicePortage.unitPrice !== undefined && Math.abs(servicePortage.unitPrice - line.tjmPortage) > 0.01) {
            validationErrors.push(`Écart Tarif B (Portage) pour ${line.serviceUser}: Facturation (${line.tjmPortage} DH) vs Services (${servicePortage.unitPrice} DH)`);
          }
        }
      });

      if (validationErrors.length > 0) {
        setError(`Blocage de sécurité : Des écarts de tarifs ont été détectés.\n\n${validationErrors.join('\n')}`);
        setProcessing(false);
        return;
      }

      // Group billing lines by client
      const groupedByClient: Record<string, BillingLine[]> = {};
      billingData.forEach(line => {
        let key = line.clientName.trim();
        if (!key && thirdPartyData.length > 0) {
          key = thirdPartyData[0].name;
        }
        if (!key) key = "Client Inconnu";

        if (!groupedByClient[key]) {
          groupedByClient[key] = [];
        }
        groupedByClient[key].push(line);
      });

      // For each client, generate a PDF
      const entries = Object.entries(groupedByClient).filter(([name]) => selectedClients.has(name));
      
      if (entries.length === 0) {
        setError("Aucun client sélectionné pour la facturation.");
        setProcessing(false);
        return;
      }

      let currentRef = invoiceReference;
      for (const [clientName, lines] of entries) {
        const clientInfo = thirdPartyData.find(tp => 
          tp.name.toLowerCase().trim() === clientName.toLowerCase().trim() ||
          tp.id.toLowerCase().trim() === clientName.toLowerCase().trim()
        ) || {
          id: 'N/A',
          name: clientName,
          address: 'Adresse non trouvée'
        };

        const invoiceLines: any[] = [];

        lines.forEach(line => {
          const nameParts = line.serviceUser.trim().split(/\s+/);
          const p1 = nameParts[0].toUpperCase();
          const p2 = nameParts[nameParts.length - 1].toUpperCase();
          const matchKey = `${p1}-${p2.charAt(0)}`;

          // Add Rate A (AO) line
          if ((billingOptions.has('GLOBAL') || billingOptions.has('RATE_A')) && line.tjmAO > 0) {
            const serviceAO = servicesData.find(s => s.ref.toUpperCase().includes('-AA-') && s.ref.toUpperCase().includes(matchKey));
            invoiceLines.push({
              service: serviceAO || { id: 'AO-MIS', ref: 'AO-MIS', description: `Tarif A - ${line.serviceUser}` },
              quantity: line.quantity,
              unitPrice: line.tjmAO,
              total: line.quantity * line.tjmAO,
              date: line.date
            });
          }

          // Add Rate B (Portage) line
          if ((billingOptions.has('GLOBAL') || billingOptions.has('RATE_B')) && line.tjmPortage > 0) {
            const servicePortage = servicesData.find(s => s.ref.toUpperCase().includes('-PR-') && s.ref.toUpperCase().includes(matchKey));
            invoiceLines.push({
              service: servicePortage || { id: 'PR-MIS', ref: 'PR-MIS', description: `Tarif B - ${line.serviceUser}` },
              quantity: line.quantity,
              unitPrice: line.tjmPortage,
              total: line.quantity * line.tjmPortage,
              date: line.date
            });
          }

          // Fallback if neither AO nor Portage but amount exists
          if (invoiceLines.length === 0 && line.amount > 0) {
             invoiceLines.push({
              service: { id: 'MIS', ref: 'MIS', description: `Prestation - ${line.serviceUser}` },
              quantity: line.quantity,
              unitPrice: line.amount / line.quantity,
              total: line.amount,
              date: line.date
            });
          }
        });

        if (invoiceLines.length > 0) {
          await createPDF(clientInfo, invoiceLines, currentRef);
          currentRef = incrementReference(currentRef);
        }
      }
      
      setInvoiceReference(currentRef);
      setError(null);
      setSuccessMessage(`Félicitations ! ${entries.length} factures ont été générées avec succès.`);
      setTimeout(() => setSuccessMessage(null), 5000);
    } catch (err) {
      console.error(err);
      setError("Erreur lors de la génération des factures. Vérifiez le format de vos fichiers.");
    } finally {
      setProcessing(false);
    }
  };

  const createPDF = async (client: any, lines: any[], invoiceRef: string) => {
    const doc = new jsPDF() as any;
    
    const invoiceDate = new Date(selectedInvoiceDate);
    const dueDate = new Date(invoiceDate);
    dueDate.setDate(dueDate.getDate() + creditDuration);

    const formatDate = (date: Date) => {
      return date.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
    };

    const totalHT = lines.reduce((sum, l) => sum + l.total, 0);
    const tva = totalHT * 0.20;
    const totalTTC = totalHT + tva;
    const totalQty = lines.reduce((sum, l) => sum + l.quantity, 0);

    const period = invoicePeriod || `${invoiceDate.getFullYear()}-${String(invoiceDate.getMonth() + 1).padStart(2, '0')}`;

    // Source Signature (Machine ID + Hardware Fingerprint)
    const machineId = getPersistentId();
    const fingerprint = getMachineFingerprint();
    const sourceData = `${machineId}|${fingerprint}|${navigator.userAgent}`;
    
    let sourceSignature = "HTTP-UNSAFE";
    let hashHex = "NO-HASH-IN-HTTP";
    
    if (typeof crypto !== 'undefined' && crypto.subtle) {
      const sourceBuffer = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(sourceData));
      const sourceHash = Array.from(new Uint8Array(sourceBuffer)).map(b => b.toString(16).padStart(2, '0')).join('');
      sourceSignature = sourceHash.substring(0, 12).toUpperCase();

      // Generate Integrity Hash (Now includes Source Signature)
      const dataToHash = `INV:${invoiceRef}|DATE:${selectedInvoiceDate}|HT:${totalHT.toFixed(2)}|TVA:${tva.toFixed(2)}|TTC:${totalTTC.toFixed(2)}|QTY:${totalQty}|SRC:${sourceSignature}${companyInfo.ice ? `|ICE:${companyInfo.ice}` : ''}`;
      const encoder = new TextEncoder();
      const data = encoder.encode(dataToHash);
      const hashBuffer = await crypto.subtle.digest('SHA-256', data);
      const hashArray = Array.from(new Uint8Array(hashBuffer));
      hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    }

    const shortHash = hashHex.substring(0, 16).toUpperCase();

    // Generate QR Code
    const dataToHash = `INV:${invoiceRef}|DATE:${selectedInvoiceDate}|HT:${totalHT.toFixed(2)}|TVA:${tva.toFixed(2)}|TTC:${totalTTC.toFixed(2)}|QTY:${totalQty}|SRC:${sourceSignature}${companyInfo.ice ? `|ICE:${companyInfo.ice}` : ''}`;
    const qrContent = `ASQSYS-VERIFY|${dataToHash}|HASH:${hashHex}`;
    const qrDataUrl = await QRCode.toDataURL(qrContent, { margin: 1, width: 100 });

    const drawHeader = (doc: any) => {
      // Color band at the very top
      doc.setFillColor(79, 70, 229); // Indigo 600
      doc.rect(0, 0, 210, 8, 'F');

      // ASQSYS & ICE in the band
      doc.setTextColor(255, 255, 255); // White
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10);
      doc.text("ASQSYS", 20, 5.5); // Left side of the band
      
      if (companyInfo.ice) {
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7);
        doc.text(`ICE: ${companyInfo.ice}`, 190, 5.5, { align: 'right' }); // Right side of the band
      }

      // INVOICE - Top Right
      doc.setFontSize(28);
      doc.setTextColor(79, 70, 229);
      doc.setFont("helvetica", "bold");
      doc.text("INVOICE", 190, 25, { align: 'right' });

      // Amount Due (MAD) - Top Right
      doc.setFontSize(10);
      doc.setTextColor(100, 116, 139); // Slate 500
      doc.setFont("helvetica", "normal");
      doc.text("Amount Due (MAD)", 190, 32, { align: 'right' });
      doc.setFontSize(16);
      doc.setTextColor(15, 23, 42); // Slate 900
      doc.setFont("helvetica", "bold");
      doc.text(formatDH(totalTTC), 190, 39, { align: 'right' });

      // QR Code - Top Right (below amount)
      doc.addImage(qrDataUrl, 'PNG', 170, 42, 20, 20);
      doc.setFontSize(6);
      doc.setTextColor(148, 163, 184);
      doc.text(`ID: ${shortHash}`, 190, 64, { align: 'right' });
      doc.text(`SRC: ${sourceSignature}`, 190, 67, { align: 'right' });

      // Invoice Details - Top Left
      doc.setFontSize(10);
      doc.setTextColor(15, 23, 42);
      doc.setFont("helvetica", "bold");
      doc.text(`Invoice Number: ${invoiceRef}`, 20, 25);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(100, 116, 139);
      doc.text(`Invoice Date: ${formatDate(invoiceDate)}`, 20, 31);
      doc.text(`Payment Due: ${formatDate(dueDate)}`, 20, 37);

      // Client Info
      doc.setFontSize(11);
      doc.setTextColor(15, 23, 42);
      doc.setFont("helvetica", "bold");
      doc.text("BILL TO", 20, 55);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10);
      doc.text(client.name, 20, 61);
      doc.setTextColor(100, 116, 139);
      doc.text(client.address || "Adresse non spécifiée", 20, 66, { maxWidth: 70 });
      let clientY = 76;
      if (client.vatNumber) {
        doc.text(`TVA: ${client.vatNumber}`, 20, clientY);
        clientY += 5;
      }
      if (client.ice) {
        doc.text(`ICE: ${client.ice}`, 20, clientY);
      }
    };

    // Table
    const tableData = lines.map(l => {
      return [
        `${l.service.description} ${period}`,
        l.quantity,
        formatDH(l.unitPrice),
        formatDH(l.total)
      ];
    });

    autoTable(doc, {
      startY: 80,
      head: [['SERVICES', 'QUANTITY', 'PRICE', 'AMOUNT']],
      body: tableData,
      theme: 'plain',
      headStyles: { 
        fillColor: [79, 70, 229], // Indigo 600
        textColor: 255,
        fontSize: 8,
        fontStyle: 'bold',
        halign: 'center'
      },
      columnStyles: {
        0: { cellWidth: 'auto' },
        1: { halign: 'center', cellWidth: 25 },
        2: { halign: 'right', cellWidth: 35 },
        3: { halign: 'right', cellWidth: 35 }
      },
      styles: { fontSize: 8, cellPadding: 1.5 },
      margin: { top: 80, bottom: 40 },
      rowPageBreak: 'avoid',
      didDrawCell: (data) => {
        if (data.section === 'body') {
          doc.setDrawColor(226, 232, 240); // Slate 200 (fine shaded line)
          doc.setLineWidth(0.1);
          doc.line(data.cell.x, data.cell.y + data.cell.height, data.cell.x + data.cell.width, data.cell.y + data.cell.height);
        }
      },
      didDrawPage: (data) => {
        // Draw Header on each page
        drawHeader(doc);

        // Footer pagination
        const pageCount = doc.internal.getNumberOfPages();
        const str = "Page " + pageCount;
        doc.setFontSize(8);
        doc.setTextColor(148, 163, 184); // Slate 400
        doc.text(str, 195, 290, { align: 'right' });
        
        // Footer Line
        doc.setDrawColor(241, 245, 249); // Slate 100
        doc.setLineWidth(0.5);
        doc.line(15, 272, 195, 272);

        // 3/4 part (Left) - Company Info
        doc.setFontSize(9);
        doc.setTextColor(15, 23, 42); // Slate 900
        doc.setFont("helvetica", "bold");
        doc.text(companyInfo.name.toUpperCase(), 15, 278);
        
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7);
        doc.setTextColor(100, 116, 139); // Slate 500
        const footerAddr = `${companyInfo.address}`;
        const footerContact = `Email: ${companyInfo.email} | TVA: ${companyInfo.vat}`;
        doc.text(footerAddr, 15, 283, { maxWidth: 135 });
        doc.text(footerContact, 15, 288);

        // 1/4 part (Right) - Slogan
        doc.setFont("helvetica", "bolditalic");
        doc.setFontSize(10);
        doc.setTextColor(79, 70, 229); // Indigo 600
        doc.text("Be Proactive Not Reactive", 195, 278, { align: 'right' });

        // Verification Note
        doc.setFontSize(6);
        doc.setTextColor(148, 163, 184);
        doc.setFont("helvetica", "normal");
        doc.text("Facture sécurisée par QR Code. Scannez pour vérifier l'intégrité des données.", 195, 283, { align: 'right' });
      }
    });

    const finalY = (doc as any).lastAutoTable.finalY;

    // Helper to check space and add page if needed
    const checkSpace = (neededHeight: number, currentYPos: number) => {
      if (currentYPos + neededHeight > 260) { // 260 is safe margin before footer line at 272
        doc.addPage();
        return 80; // Start below header
      }
      return currentYPos;
    };

    // Totals Box
    let currentY = checkSpace(50, finalY);
    const boxWidth = 70;
    const startX = 190 - boxWidth;
    
    doc.setFontSize(10);
    doc.setTextColor(100, 116, 139); // Slate 500
    doc.text(`Total HT:`, startX, currentY + 15);
    doc.text(formatDH(totalHT), 190, currentY + 15, { align: 'right' });
    
    doc.text(`TVA (20%):`, startX, currentY + 22);
    doc.text(formatDH(tva), 190, currentY + 22, { align: 'right' });
    
    doc.setDrawColor(79, 70, 229); // Indigo 600
    doc.setLineWidth(0.5);
    doc.line(startX, currentY + 26, 190, currentY + 26);
    
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.setTextColor(15, 23, 42); // Slate 900
    doc.text(`TOTAL TTC:`, startX, currentY + 33);
    doc.text(formatDH(totalTTC), 190, currentY + 33, { align: 'right' });

    // Move currentY to after the Totals Box
    currentY += 40;

    // --- OPTIMIZED FOOTER BLOCKS (Grouped to prevent splitting) ---
    const footerPadding = 4;
    const amountInWordsHeight = 14;
    
    // Calculate Bank height
    const hasBankInfo = (selectedBankAccounts.account1 && companyInfo.bankAccount1) || (selectedBankAccounts.account2 && companyInfo.bankAccount2);
    const hasNote = invoiceNote.trim().length > 0;
    const noteWidth = hasNote && hasBankInfo ? 100 : 180;
    const noteLines = hasNote ? doc.splitTextToSize(`Note : ${invoiceNote}`, noteWidth - 10) : [];
    const noteHeight = hasNote ? (noteLines.length * 4) + 8 : 0;

    const bankWidth = hasNote ? 75 : 180;
    const bankAccountCount = (selectedBankAccounts.account1 && companyInfo.bankAccount1 ? 1 : 0) + (selectedBankAccounts.account2 && companyInfo.bankAccount2 ? 1 : 0);
    const bankHeight = hasBankInfo ? 22 + (bankAccountCount * 5) : 0;

    const footerGroupHeight = amountInWordsHeight + 5 + Math.max(noteHeight, bankHeight);
    
    // Ensure the entire block fits on the same page
    currentY = checkSpace(footerGroupHeight + 10, currentY);

    // 1. Amount in Words (Full Width, Compact)
    doc.setFillColor(248, 250, 252); // Slate 50
    doc.roundedRect(15, currentY, 180, amountInWordsHeight, 2, 2, 'F');
    doc.setFontSize(7.5);
    doc.setTextColor(100, 116, 139); // Slate 500
    doc.setFont("helvetica", "italic");
    doc.text(`Arrêté la présente facture à la somme de :`, 20, currentY + 5);
    doc.setFont("helvetica", "bolditalic");
    doc.setTextColor(15, 23, 42); // Slate 900
    doc.text(`*** ${numberToWordsFR(totalTTC)} ***`, 20, currentY + 10, { maxWidth: 170 });

    currentY += amountInWordsHeight + 4;

    // 2. Side-by-Side: Note & Bank Details
    const startY = currentY;

    if (hasNote) {
      doc.setFillColor(255, 251, 235); // Amber 50
      doc.setDrawColor(251, 191, 36); // Amber 400
      doc.setLineWidth(0.1);
      doc.roundedRect(15, startY, noteWidth, noteHeight, 2, 2, 'FD');
      
      doc.setFontSize(7.5);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(146, 64, 14); // Amber 800
      doc.text(noteLines, 20, startY + 6);
    }

    if (hasBankInfo) {
      const bankX = hasNote ? 15 + noteWidth + 5 : 15;
      doc.setFillColor(241, 245, 249); // Slate 100
      doc.roundedRect(bankX, startY, bankWidth, bankHeight, 2, 2, 'F');
      
      doc.setFontSize(8);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(15, 23, 42);
      doc.text("MODE DE PAIEMENT", bankX + 5, startY + 6);
      
      doc.setFontSize(7);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(71, 85, 105); // Slate 600
      doc.text("Virement bancaire vers :", bankX + 5, startY + 11);
      
      let bankY = startY + 16;
      doc.setFont("helvetica", "bold");
      doc.setTextColor(15, 23, 42);
      
      if (selectedBankAccounts.account1 && companyInfo.bankAccount1) {
        doc.text(companyInfo.bankAccount1, bankX + 5, bankY);
        bankY += 5;
      }
      if (selectedBankAccounts.account2 && companyInfo.bankAccount2) {
        doc.text(companyInfo.bankAccount2, bankX + 5, bankY);
        bankY += 5;
      }

      // Payment reference reminder (Integrated into bank box)
      doc.setFontSize(7);
      doc.setFont("helvetica", "bolditalic");
      doc.setTextColor(79, 70, 229); // Indigo 600
      doc.text(`IMPORTANT : Rappeler la réf. ${invoiceRef}`, bankX + 5, bankY + 2);
    }

    const fileName = `INV-${invoiceDate.getFullYear()}-${String(invoiceDate.getMonth() + 1).padStart(2, '0')}-ASQSYS-${client.name.replace(/[^a-z0-9]/gi, '_')}-${invoiceRef}.pdf`;
    doc.save(fileName);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-4 md:p-12">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <header className="mb-16 flex flex-col md:flex-row items-center justify-between gap-8">
          <div className="text-center md:text-left">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              className="inline-flex items-center justify-center p-4 bg-indigo-600 rounded-3xl mb-6 shadow-xl shadow-indigo-100"
            >
              <Calculator className="w-10 h-10 text-white" />
            </motion.div>
            <h1 className="text-5xl font-black tracking-tight text-slate-900 mb-3">
              ASQ<span className="text-indigo-600">SYS</span>
            </h1>
            <p className="text-slate-500 text-xl font-medium">Générez vos factures PDF professionnelles en un clic</p>
          </div>
          <div className="flex flex-wrap items-center justify-center md:justify-end gap-4">
            <button 
              onClick={() => setShowSettings(true)}
              className="flex items-center space-x-3 px-8 py-4 bg-white border border-slate-200 rounded-3xl font-bold text-slate-600 hover:bg-slate-50 hover:border-slate-300 transition-all shadow-sm card-shadow"
            >
              <Building2 className="w-6 h-6 text-indigo-600" />
              <span>Ma Société</span>
            </button>
            
            <button 
              onClick={() => setShowFAQ(true)}
              className="flex items-center space-x-3 px-8 py-4 bg-indigo-50 border border-indigo-100 rounded-3xl font-bold text-indigo-600 hover:bg-indigo-100 transition-all shadow-sm"
            >
              <HelpCircle className="w-6 h-6" />
              <span>Guide (FAQ)</span>
            </button>
          </div>
        </header>

        {/* Main Content */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 mb-12">
          <FileCard 
            title="Ma Société" 
            icon={<Building2 className="w-5 h-5" />}
            uploaded={filesUploaded.company}
            count={companies.length}
            onChange={(e) => handleFileUpload(e, 'company')}
            description="Infos émetteur"
            unit="entités"
          />
          <FileCard 
            title="Clients" 
            icon={<Users className="w-5 h-5" />}
            uploaded={filesUploaded.thirdParty}
            count={thirdPartyData.length}
            onChange={(e) => handleFileUpload(e, 'thirdParty')}
            description="Liste des tiers"
          />
          <FileCard 
            title="Services" 
            icon={<Package className="w-5 h-5" />}
            uploaded={filesUploaded.services}
            count={servicesData.length}
            onChange={(e) => handleFileUpload(e, 'services')}
            description="Catalogue prestations"
          />
          <FileCard 
            title="Facturation" 
            icon={<FileSpreadsheet className="w-5 h-5" />}
            uploaded={filesUploaded.billing}
            count={billingData.length}
            onChange={(e) => handleFileUpload(e, 'billing')}
            description="Données d'activité"
          />
        </div>

        {/* Client Selection Section */}
        {detectedClients.length > 0 && (
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white rounded-[2rem] p-8 shadow-sm border border-slate-100 mb-8 card-shadow"
          >
            <div className="flex items-center justify-between mb-6">
              <div>
                <h2 className="text-xl font-black text-slate-800">Clients détectés</h2>
                <p className="text-sm text-slate-400 font-medium">Sélectionnez les clients à facturer</p>
              </div>
              <div className="flex items-center space-x-3">
                <button 
                  onClick={() => setSelectedClients(new Set(detectedClients))}
                  className="text-xs font-black text-indigo-600 uppercase tracking-widest hover:text-indigo-700 transition-colors"
                >
                  Tout cocher
                </button>
                <div className="w-px h-3 bg-slate-200" />
                <button 
                  onClick={() => setSelectedClients(new Set())}
                  className="text-xs font-black text-slate-400 uppercase tracking-widest hover:text-slate-500 transition-colors"
                >
                  Tout décocher
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {detectedClients.map(client => (
                <label 
                  key={client}
                  className={cn(
                    "flex items-center space-x-3 p-4 rounded-2xl border cursor-pointer transition-all duration-300",
                    selectedClients.has(client) 
                      ? "bg-indigo-50 border-indigo-100 text-indigo-700" 
                      : "bg-white border-slate-100 text-slate-400 hover:border-slate-200"
                  )}
                >
                  <div className={cn(
                    "w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-all",
                    selectedClients.has(client)
                      ? "bg-indigo-600 border-indigo-600"
                      : "bg-white border-slate-200"
                  )}>
                    {selectedClients.has(client) && <Check className="w-3 h-3 text-white" />}
                  </div>
                  <input 
                    type="checkbox"
                    className="hidden"
                    checked={selectedClients.has(client)}
                    onChange={() => {
                      const next = new Set(selectedClients);
                      if (next.has(client)) next.delete(client);
                      else next.add(client);
                      setSelectedClients(next);
                    }}
                  />
                  <span className="font-bold truncate">{client}</span>
                </label>
              ))}
            </div>
          </motion.div>
        )}

        {/* Action Section */}
        <div className="bg-white rounded-[2rem] p-10 shadow-sm border border-slate-100 mb-12 card-shadow">
          <div className="flex flex-col items-center justify-center space-y-8">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-8 w-full max-w-6xl">
              <div>
                <label className="block text-[10px] font-black text-slate-400 mb-3 uppercase tracking-[0.2em] ml-1">
                  Date de facturation
                </label>
                <select 
                  value={selectedInvoiceDate}
                  onChange={(e) => setSelectedInvoiceDate(e.target.value)}
                  className="w-full px-5 py-4 bg-slate-50 border border-slate-100 rounded-2xl font-bold text-slate-700 focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all appearance-none cursor-pointer"
                >
                  {getLastDaysOfMonths().map(date => (
                    <option key={date} value={date}>
                      {new Date(date).toLocaleDateString('fr-FR', { month: 'long', year: 'numeric', day: 'numeric' })}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-[10px] font-black text-slate-400 mb-3 uppercase tracking-[0.2em] ml-1">
                  Période (Libellé)
                </label>
                <input 
                  type="text"
                  value={invoicePeriod}
                  onChange={(e) => setInvoicePeriod(e.target.value)}
                  placeholder="Ex: 2026-01"
                  className="w-full px-5 py-4 bg-slate-50 border border-slate-100 rounded-2xl font-bold text-slate-700 focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all"
                />
              </div>
              <div>
                <label className="block text-[10px] font-black text-slate-400 mb-3 uppercase tracking-[0.2em] ml-1">
                  Réf. de départ
                </label>
                <input 
                  type="text"
                  value={invoiceReference}
                  onChange={(e) => setInvoiceReference(e.target.value)}
                  placeholder="Ex: FAC-2026-001"
                  className="w-full px-5 py-4 bg-slate-50 border border-slate-100 rounded-2xl font-bold text-slate-700 focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all"
                />
              </div>
              <div className="lg:col-span-2">
                <label className="block text-[10px] font-black text-slate-400 mb-3 uppercase tracking-[0.2em] ml-1">
                  Type de facturation (Mode d'application des tarifs)
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <label className={cn(
                    "flex items-center space-x-3 p-4 rounded-2xl border cursor-pointer transition-all",
                    billingOptions.has('GLOBAL') ? "bg-indigo-50 border-indigo-200 text-indigo-700" : "bg-slate-50 border-slate-100 text-slate-400 hover:border-slate-200"
                  )}>
                    <input 
                      type="checkbox" 
                      className="hidden"
                      checked={billingOptions.has('GLOBAL')}
                      onChange={() => {
                        setBillingOptions(new Set(['GLOBAL']));
                      }}
                    />
                    <div className={cn(
                      "w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-all",
                      billingOptions.has('GLOBAL') ? "bg-indigo-600 border-indigo-600" : "bg-white border-slate-200"
                    )}>
                      {billingOptions.has('GLOBAL') && <Check className="w-3 h-3 text-white" />}
                    </div>
                    <div className="flex flex-col">
                      <span className="font-bold text-xs">Mode Global</span>
                      <span className="text-[9px] opacity-70">Consolidé (ex: TJM Global, Package)</span>
                    </div>
                  </label>

                  <label className={cn(
                    "flex items-center space-x-3 p-4 rounded-2xl border cursor-pointer transition-all",
                    billingOptions.has('RATE_A') ? "bg-indigo-50 border-indigo-200 text-indigo-700" : "bg-slate-50 border-slate-100 text-slate-400 hover:border-slate-200"
                  )}>
                    <input 
                      type="checkbox" 
                      className="hidden"
                      checked={billingOptions.has('RATE_A')}
                      onChange={() => {
                        const next = new Set(billingOptions);
                        next.delete('GLOBAL');
                        if (next.has('RATE_A')) {
                          next.delete('RATE_A');
                          if (next.size === 0) next.add('GLOBAL');
                        } else {
                          next.add('RATE_A');
                        }
                        setBillingOptions(next);
                      }}
                    />
                    <div className={cn(
                      "w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-all",
                      billingOptions.has('RATE_A') ? "bg-indigo-600 border-indigo-600" : "bg-white border-slate-200"
                    )}>
                      {billingOptions.has('RATE_A') && <Check className="w-3 h-3 text-white" />}
                    </div>
                    <div className="flex flex-col">
                      <span className="font-bold text-xs">Tarif Type A</span>
                      <span className="text-[9px] opacity-70">ex: TJM A.O, Prix VIP, Vente</span>
                    </div>
                  </label>

                  <label className={cn(
                    "flex items-center space-x-3 p-4 rounded-2xl border cursor-pointer transition-all",
                    billingOptions.has('RATE_B') ? "bg-indigo-50 border-indigo-200 text-indigo-700" : "bg-slate-50 border-slate-100 text-slate-400 hover:border-slate-200"
                  )}>
                    <input 
                      type="checkbox" 
                      className="hidden"
                      checked={billingOptions.has('RATE_B')}
                      onChange={() => {
                        const next = new Set(billingOptions);
                        next.delete('GLOBAL');
                        if (next.has('RATE_B')) {
                          next.delete('RATE_B');
                          if (next.size === 0) next.add('GLOBAL');
                        } else {
                          next.add('RATE_B');
                        }
                        setBillingOptions(next);
                      }}
                    />
                    <div className={cn(
                      "w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-all",
                      billingOptions.has('RATE_B') ? "bg-indigo-600 border-indigo-600" : "bg-white border-slate-200"
                    )}>
                      {billingOptions.has('RATE_B') && <Check className="w-3 h-3 text-white" />}
                    </div>
                    <div className="flex flex-col">
                      <span className="font-bold text-xs">Tarif Type B</span>
                      <span className="text-[9px] opacity-70">ex: TJM Portage, Normal, Achat</span>
                    </div>
                  </label>
                </div>
              </div>
              <div>
                <label className="block text-[10px] font-black text-slate-400 mb-3 uppercase tracking-[0.2em] ml-1">
                  Durée crédit (jours)
                </label>
                <input 
                  type="number"
                  value={creditDuration}
                  onChange={(e) => setCreditDuration(parseInt(e.target.value) || 0)}
                  className="w-full px-5 py-4 bg-slate-50 border border-slate-100 rounded-2xl font-bold text-slate-700 focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all"
                />
              </div>
              <div className="sm:col-span-3">
                <label className="block text-[10px] font-black text-slate-400 mb-3 uppercase tracking-[0.2em] ml-1">
                  Note (Optionnel - Apparaît sur la facture)
                </label>
                <textarea 
                  value={invoiceNote}
                  onChange={(e) => setInvoiceNote(e.target.value)}
                  placeholder="Notes explicatives pour le comptable ou précisions sur la prestation..."
                  className="w-full px-5 py-4 bg-slate-50 border border-slate-100 rounded-2xl font-medium text-slate-700 focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all resize-none h-24"
                />
              </div>
            </div>

            {/* Bank Account Selection */}
            {filesUploaded.company && (companyInfo.bankAccount1 || companyInfo.bankAccount2) && (
              <div className="w-full max-w-4xl">
                <label className="block text-[10px] font-black text-slate-400 mb-3 uppercase tracking-[0.2em] ml-1">
                  Comptes bancaires à afficher (Obligatoire)
                </label>
                <div className="flex flex-col sm:flex-row gap-4">
                  {companyInfo.bankAccount1 && (
                    <label className={cn(
                      "flex-1 flex items-center space-x-3 p-4 rounded-2xl border cursor-pointer transition-all",
                      selectedBankAccounts.account1 ? "bg-indigo-50 border-indigo-100 text-indigo-700" : "bg-slate-50 border-slate-100 text-slate-400"
                    )}>
                      <div className={cn(
                        "w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-all",
                        selectedBankAccounts.account1 ? "bg-indigo-600 border-indigo-600" : "bg-white border-slate-200"
                      )}>
                        {selectedBankAccounts.account1 && <Check className="w-3 h-3 text-white" />}
                      </div>
                      <input 
                        type="checkbox" 
                        className="hidden"
                        checked={selectedBankAccounts.account1}
                        onChange={(e) => setSelectedBankAccounts(prev => ({ ...prev, account1: e.target.checked }))}
                      />
                      <span className="font-bold text-sm truncate">{companyInfo.bankAccount1}</span>
                    </label>
                  )}
                  {companyInfo.bankAccount2 && (
                    <label className={cn(
                      "flex-1 flex items-center space-x-3 p-4 rounded-2xl border cursor-pointer transition-all",
                      selectedBankAccounts.account2 ? "bg-indigo-50 border-indigo-100 text-indigo-700" : "bg-slate-50 border-slate-100 text-slate-400"
                    )}>
                      <div className={cn(
                        "w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-all",
                        selectedBankAccounts.account2 ? "bg-indigo-600 border-indigo-600" : "bg-white border-slate-200"
                      )}>
                        {selectedBankAccounts.account2 && <Check className="w-3 h-3 text-white" />}
                      </div>
                      <input 
                        type="checkbox" 
                        className="hidden"
                        checked={selectedBankAccounts.account2}
                        onChange={(e) => setSelectedBankAccounts(prev => ({ ...prev, account2: e.target.checked }))}
                      />
                      <span className="font-bold text-sm truncate">{companyInfo.bankAccount2}</span>
                    </label>
                  )}
                </div>
              </div>
            )}

            <div className="flex items-center space-x-4 flex-wrap justify-center gap-y-4">
              <StatusBadge active={filesUploaded.company} label="Société" />
              <div className="hidden sm:block w-8 h-px bg-slate-100" />
              <StatusBadge active={filesUploaded.thirdParty} label="Tiers" />
              <div className="hidden sm:block w-8 h-px bg-slate-100" />
              <StatusBadge active={filesUploaded.services} label="Services" />
              <div className="hidden sm:block w-8 h-px bg-slate-100" />
              <StatusBadge active={filesUploaded.billing} label="Facturation" />
            </div>

            {error && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="flex items-center space-x-2 text-red-500 bg-red-50 px-6 py-3 rounded-2xl border border-red-100"
              >
                <AlertCircle className="w-5 h-5" />
                <span className="text-sm font-bold">{error}</span>
              </motion.div>
            )}

            {successMessage && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="flex items-center space-x-2 bg-emerald-50 text-emerald-600 px-8 py-5 rounded-[2rem] border border-emerald-100 mb-6"
              >
                <CheckCircle2 className="w-6 h-6" />
                <span className="text-base font-black">{successMessage}</span>
              </motion.div>
            )}

            <button
              onClick={generateInvoices}
              disabled={!filesUploaded.company || !filesUploaded.thirdParty || !filesUploaded.services || !filesUploaded.billing || processing}
              className={cn(
                "group relative flex items-center justify-center space-x-4 px-16 py-6 rounded-[2rem] font-black text-2xl transition-all duration-500 shadow-2xl",
                filesUploaded.company && filesUploaded.thirdParty && filesUploaded.services && filesUploaded.billing && !processing
                  ? "bg-indigo-600 text-white hover:bg-indigo-700 hover:scale-[1.02] active:scale-[0.98] shadow-indigo-200"
                  : "bg-slate-100 text-slate-300 cursor-not-allowed shadow-none"
              )}
            >
              {processing ? (
                <div className="flex items-center space-x-3">
                  <div className="w-8 h-8 border-4 border-white/30 border-t-white rounded-full animate-spin" />
                  <span>Génération...</span>
                </div>
              ) : (
                <>
                  <Download className="w-6 h-6" />
                  <span>Générer les Factures</span>
                </>
              )}
            </button>
            
            <p className="text-sm text-slate-400 font-bold tracking-tight">
              {processing ? "Traitement en cours..." : "Tous les fichiers doivent être chargés pour commencer"}
            </p>
          </div>
        </div>

        {/* Data Preview */}
        {(billingData.length > 0) && (
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white rounded-3xl p-8 shadow-sm border border-gray-100 overflow-hidden mb-8"
          >
            <h2 className="text-xl font-bold mb-6 flex items-center gap-2 text-slate-800">
              <FileText className="w-5 h-5 text-indigo-600" />
              Aperçu des données à facturer
            </h2>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm">
                <thead>
                  <tr className="border-b border-slate-100">
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px]">Client</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px]">Collaborateur</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-right">Nbre Jours</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-right">TJM A.O</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-right">TJM Portage</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-right">Total HT</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-center">Match</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {billingData.slice(0, 5).map((line, i) => {
                    const nameParts = line.serviceUser.trim().split(/\s+/);
                    let isMatched = false;
                    if (nameParts.length >= 2) {
                      const p1 = nameParts[0].toUpperCase();
                      const p2 = nameParts[nameParts.length - 1].toUpperCase();
                      const key = `${p1}-${p2.charAt(0)}`;
                      isMatched = servicesData.some(s => s.ref.toUpperCase().includes(key));
                    }
                    const totalHT = line.amount || (line.quantity * (line.tjmAO + line.tjmPortage));
                    return (
                      <tr key={i} className="hover:bg-gray-50/50 transition-colors">
                        <td className="py-4 font-bold text-indigo-600">{line.clientName}</td>
                        <td className="py-4 text-slate-600 font-medium">{line.serviceUser}</td>
                        <td className="py-4 text-right font-mono text-slate-500">{line.quantity}</td>
                        <td className="py-4 text-right font-mono text-slate-500">{formatDH(line.tjmAO)}</td>
                        <td className="py-4 text-right font-mono text-slate-500">{formatDH(line.tjmPortage)}</td>
                        <td className="py-4 text-right font-bold text-slate-700">{formatDH(totalHT)}</td>
                        <td className="py-4 text-center">
                          {isMatched ? (
                            <span className="px-2 py-1 bg-green-50 text-green-600 rounded-full text-[10px] font-bold uppercase tracking-wider">OK</span>
                          ) : (
                            <span className="px-2 py-1 bg-red-50 text-red-600 rounded-full text-[10px] font-bold uppercase tracking-wider">KO</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {billingData.length > 5 && (
                <p className="mt-4 text-center text-gray-400 text-xs italic">
                  + {billingData.length - 5} autres lignes...
                </p>
              )}
            </div>
          </motion.div>
        )}

        {/* Services Preview */}
        {(servicesData.length > 0) && (
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white rounded-3xl p-8 shadow-sm border border-gray-100 overflow-hidden mb-8"
          >
            <h2 className="text-xl font-bold mb-6 flex items-center gap-2">
              <Package className="w-5 h-5 text-blue-600" />
              Catalogue Services ({servicesData.length})
            </h2>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm">
                <thead>
                  <tr className="border-b border-gray-100">
                    <th className="pb-4 font-bold text-gray-400 uppercase tracking-wider text-xs">ID (Facture)</th>
                    <th className="pb-4 font-bold text-gray-400 uppercase tracking-wider text-xs">Réf. (Liaison)</th>
                    <th className="pb-4 font-bold text-gray-400 uppercase tracking-wider text-xs">Description</th>
                    <th className="pb-4 font-bold text-gray-400 uppercase tracking-wider text-xs text-right">TJM (Services)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {servicesData.slice(0, 5).map((service, i) => {
                    return (
                      <tr key={i} className="hover:bg-gray-50/50 transition-colors">
                        <td className="py-4 font-mono text-blue-600 text-xs">{service.id}</td>
                        <td className="py-4 font-mono text-gray-500 text-xs">{service.ref}</td>
                        <td className="py-4 text-gray-600">{service.description} {invoicePeriod}</td>
                        <td className="py-4 text-right font-mono text-slate-500">{service.unitPrice ? formatDH(service.unitPrice) : '-'}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </motion.div>
        )}

        {/* FAQ Modal */}
        <AnimatePresence>
          {showFAQ && (
            <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 md:p-8">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setShowFAQ(false)}
                className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              />
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 20 }}
                className="relative w-full max-w-4xl max-h-[90vh] bg-white rounded-[2.5rem] shadow-2xl overflow-hidden flex flex-col"
              >
                <div className="p-8 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                  <div className="flex items-center space-x-4">
                    <div className="p-3 bg-indigo-600 rounded-2xl">
                      <HelpCircle className="w-6 h-6 text-white" />
                    </div>
                    <div>
                      <h2 className="text-2xl font-black text-slate-900">Guide d'utilisation</h2>
                      <p className="text-slate-500 font-medium">Structure des fichiers et conseils</p>
                    </div>
                  </div>
                  <button 
                    onClick={() => setShowFAQ(false)}
                    className="p-3 hover:bg-slate-100 rounded-2xl transition-colors"
                  >
                    <X className="w-6 h-6 text-slate-400" />
                  </button>
                </div>

                <div className="p-8 overflow-y-auto custom-scrollbar space-y-12">
                  {/* Introduction */}
                  <section>
                    <h3 className="text-lg font-black text-indigo-600 uppercase tracking-widest mb-4">Introduction</h3>
                    <p className="text-slate-600 leading-relaxed">
                      Pour générer vos factures, l'application nécessite 4 fichiers Excel spécifiques. 
                      L'application est flexible sur les noms de colonnes (elle cherche des mots-clés), 
                      mais il est recommandé de suivre les structures ci-dessous pour une détection optimale.
                    </p>
                  </section>

                  {/* 1. Fichier Société */}
                  <section className="bg-slate-50 p-8 rounded-[2rem] border border-slate-100">
                    <div className="flex items-center space-x-3 mb-6">
                      <Building2 className="w-6 h-6 text-indigo-600" />
                      <h3 className="text-xl font-black text-slate-900">1. Fichier Société (Ma Société)</h3>
                    </div>
                    <p className="text-sm text-slate-500 mb-4 italic">Contient vos informations légales et bancaires.</p>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="space-y-2">
                        <p className="text-xs font-black text-slate-400 uppercase tracking-wider">Colonnes obligatoires</p>
                        <ul className="text-sm text-slate-700 space-y-1">
                          <li>• <span className="font-bold">Name / Nom</span> : Votre raison sociale</li>
                          <li>• <span className="font-bold">Address / Adresse</span> : Siège social</li>
                          <li>• <span className="font-bold">Email</span> : Contact facturation</li>
                          <li>• <span className="font-bold">VAT ID / TVA</span> : Votre numéro de TVA</li>
                        </ul>
                      </div>
                      <div className="space-y-2">
                        <p className="text-xs font-black text-slate-400 uppercase tracking-wider">Colonnes de paiement</p>
                        <ul className="text-sm text-slate-700 space-y-1">
                          <li>• <span className="font-bold">Payment bank account 1</span> : IBAN/RIB principal</li>
                          <li>• <span className="font-bold">Payment bank account 2</span> : IBAN/RIB secondaire</li>
                        </ul>
                      </div>
                    </div>
                  </section>

                  {/* 2. Fichier Tiers */}
                  <section className="bg-slate-50 p-8 rounded-[2rem] border border-slate-100">
                    <div className="flex items-center space-x-3 mb-6">
                      <Users className="w-6 h-6 text-indigo-600" />
                      <h3 className="text-xl font-black text-slate-900">2. Fichier Tiers (Clients)</h3>
                    </div>
                    <p className="text-sm text-slate-500 mb-4 italic">Base de données de vos clients.</p>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="space-y-2">
                        <p className="text-xs font-black text-slate-400 uppercase tracking-wider">Colonnes clés</p>
                        <ul className="text-sm text-slate-700 space-y-1">
                          <li>• <span className="font-bold">Name / Nom</span> : Nom du client (doit correspondre au fichier facturation)</li>
                          <li>• <span className="font-bold">Address / Adresse</span> : Adresse de facturation</li>
                          <li>• <span className="font-bold">VAT ID / TVA</span> : Numéro de TVA du client</li>
                        </ul>
                      </div>
                      <div className="space-y-2">
                        <p className="text-xs font-black text-slate-400 uppercase tracking-wider">Optionnel</p>
                        <ul className="text-sm text-slate-700 space-y-1">
                          <li>• <span className="font-bold">Customer Code</span> : Votre référence interne</li>
                          <li>• <span className="font-bold">Email</span> : Pour vos archives</li>
                        </ul>
                      </div>
                    </div>
                  </section>

                  {/* 3. Fichier Prestations */}
                  <section className="bg-slate-50 p-8 rounded-[2rem] border border-slate-100">
                    <div className="flex items-center space-x-3 mb-6">
                      <Package className="w-6 h-6 text-indigo-600" />
                      <h3 className="text-xl font-black text-slate-900">3. Fichier Prestations (Services)</h3>
                    </div>
                    <p className="text-sm text-slate-500 mb-4 italic">Catalogue de vos services pour le rapprochement automatique.</p>
                    <ul className="text-sm text-slate-700 space-y-2">
                      <li>• <span className="font-bold">Ref / Reference</span> : Code unique du service (ex: CONS-DEV). <br/><span className="text-xs text-indigo-500">Note: L'application tente de faire correspondre ce code avec le nom de l'intervenant dans le fichier facturation.</span></li>
                      <li>• <span className="font-bold">Label / Description</span> : Libellé qui apparaîtra sur la facture.</li>
                    </ul>
                  </section>

                  {/* 4. Fichier Facturation */}
                  <section className="bg-indigo-600 p-8 rounded-[2rem] text-white">
                    <div className="flex items-center space-x-3 mb-6">
                      <FileSpreadsheet className="w-6 h-6 text-white" />
                      <h3 className="text-xl font-black">4. Fichier Facturation (Données)</h3>
                    </div>
                    <p className="text-sm text-indigo-100 mb-6 italic">C'est le fichier moteur qui contient les lignes de prestations du mois.</p>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-3">
                        <p className="text-xs font-black text-indigo-200 uppercase tracking-wider">Colonnes de calcul</p>
                        <ul className="text-sm space-y-2">
                          <li>• <span className="font-bold">Nbre de jours / Quantity</span> : Quantité (ex: 15)</li>
                          <li>• <span className="font-bold">TJM A.O</span> : Tarif Journalier Moyen A.O</li>
                          <li>• <span className="font-bold">TJM Portage</span> : Tarif Journalier Moyen Portage</li>
                          <li>• <span className="font-bold">Total HT / Amount</span> : Montant total de la ligne</li>
                        </ul>
                      </div>
                      <div className="space-y-3">
                        <p className="text-xs font-black text-indigo-200 uppercase tracking-wider">Colonnes de liaison</p>
                        <ul className="text-sm space-y-2">
                          <li>• <span className="font-bold">Client Name</span> : Doit correspondre au nom dans le fichier Tiers.</li>
                          <li>• <span className="font-bold">Service User</span> : Nom de l'intervenant (ex: NAOUAR M).</li>
                          <li>• <span className="font-bold">Date</span> : Date de la prestation.</li>
                        </ul>
                      </div>
                    </div>
                    <div className="mt-6 p-4 bg-white/10 rounded-xl border border-white/20">
                      <p className="text-xs font-bold mb-2 uppercase tracking-widest opacity-80">Règle de matching</p>
                      <p className="text-sm leading-relaxed">
                        L'application cherche dans le fichier <span className="font-bold underline">Services</span> une référence contenant :<br/>
                        1. <span className="font-bold">-AA-</span> pour le TJM A.O<br/>
                        2. <span className="font-bold">-PR-</span> pour le TJM Portage<br/>
                        3. Le code intervenant (ex: <span className="font-bold">NAOUAR-M</span>)
                      </p>
                    </div>
                  </section>

                  {/* Conseils */}
                  <section className="p-8 border-2 border-dashed border-slate-100 rounded-[2rem]">
                    <h3 className="text-lg font-black text-slate-900 mb-4">💡 Conseils d'expert</h3>
                    <ul className="text-sm text-slate-600 space-y-3">
                      <li className="flex items-start space-x-3">
                        <div className="w-1.5 h-1.5 rounded-full bg-indigo-600 mt-1.5 shrink-0" />
                        <span><span className="font-bold text-slate-900">Rapprochement intelligent :</span> L'application utilise le "Service User" pour trouver la prestation correspondante. Assurez-vous que les noms d'intervenants sont cohérents.</span>
                      </li>
                      <li className="flex items-start space-x-3">
                        <div className="w-1.5 h-1.5 rounded-full bg-indigo-600 mt-1.5 shrink-0" />
                        <span><span className="font-bold text-slate-900">Noms de colonnes :</span> Si une colonne n'est pas détectée, essayez d'utiliser les termes exacts cités ci-dessus (en français ou anglais).</span>
                      </li>
                      <li className="flex items-start space-x-3">
                        <div className="w-1.5 h-1.5 rounded-full bg-indigo-600 mt-1.5 shrink-0" />
                        <span><span className="font-bold text-slate-900">Sécurité :</span> Vos fichiers sont traités localement dans votre navigateur. Aucune donnée n'est envoyée sur un serveur externe.</span>
                      </li>
                    </ul>
                  </section>
                </div>

                <div className="p-8 bg-slate-50 border-t border-slate-100 flex justify-center">
                  <button 
                    onClick={() => setShowFAQ(false)}
                    className="px-12 py-4 bg-indigo-600 text-white font-black rounded-2xl hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200"
                  >
                    J'ai compris
                  </button>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Settings Modal */}
        <AnimatePresence>
          {showSettings && (
            <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 md:p-8">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setShowSettings(false)}
                className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              />
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 20 }}
                className="relative w-full max-w-2xl max-h-[90vh] bg-white rounded-[2.5rem] shadow-2xl overflow-hidden flex flex-col"
              >
                <div className="p-8 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                  <div className="flex items-center space-x-4">
                    <div className="p-3 bg-indigo-600 rounded-2xl">
                      <Building2 className="w-6 h-6 text-white" />
                    </div>
                    <div>
                      <h2 className="text-2xl font-black text-slate-900">Ma Société</h2>
                      <p className="text-slate-500 font-medium">Gérez vos entités de facturation</p>
                    </div>
                  </div>
                  <button 
                    onClick={() => setShowSettings(false)}
                    className="p-3 hover:bg-slate-100 rounded-2xl transition-colors"
                  >
                    <X className="w-6 h-6 text-slate-400" />
                  </button>
                </div>
                
                <div className="p-8 overflow-y-auto custom-scrollbar space-y-8">
                  {/* Company Selection List */}
                  <section>
                    <div className="flex items-center justify-between mb-4">
                      <h3 className="text-xs font-black text-slate-400 uppercase tracking-widest">Vos Sociétés ({companies.length})</h3>
                      <button 
                        onClick={() => document.getElementById('company-upload-settings')?.click()}
                        className="text-[10px] font-black text-indigo-600 uppercase tracking-widest hover:text-indigo-700 transition-colors flex items-center gap-1"
                      >
                        <Upload className="w-3 h-3" />
                        Ajouter
                      </button>
                      <input 
                        id="company-upload-settings"
                        type="file"
                        className="hidden"
                        onChange={(e) => handleFileUpload(e, 'company')}
                        accept=".xlsx,.xls"
                      />
                    </div>
                    <div className="grid grid-cols-1 gap-3">
                      {companies.map((company) => (
                        <div 
                          key={company.id}
                          onClick={() => setSelectedCompanyId(company.id)}
                          className={cn(
                            "group flex items-center justify-between p-5 rounded-2xl border-2 transition-all cursor-pointer",
                            selectedCompanyId === company.id 
                              ? "border-indigo-600 bg-indigo-50/50 shadow-lg shadow-indigo-100" 
                              : "border-slate-100 hover:border-slate-200 bg-white"
                          )}
                        >
                          <div className="flex items-center space-x-4">
                            <div className={cn(
                              "w-12 h-12 rounded-xl flex items-center justify-center font-black text-lg",
                              selectedCompanyId === company.id ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-400"
                            )}>
                              {company.name.charAt(0).toUpperCase()}
                            </div>
                            <div>
                              <p className="font-black text-slate-900">{company.name}</p>
                              <p className="text-xs text-slate-500 font-medium truncate max-w-[200px]">{company.address}</p>
                            </div>
                          </div>
                          
                          <div className="flex items-center space-x-2">
                            {selectedCompanyId === company.id && (
                              <div className="px-3 py-1 bg-indigo-600 text-white text-[10px] font-black rounded-full uppercase tracking-wider">Actif</div>
                            )}
                            <button 
                              onClick={(e) => {
                                e.stopPropagation();
                                if (confirm(`Supprimer la société ${company.name} ?`)) {
                                  setCompanies(prev => prev.filter(c => c.id !== company.id));
                                  if (selectedCompanyId === company.id) setSelectedCompanyId(null);
                                }
                              }}
                              className="p-2 text-slate-300 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          </div>
                        </div>
                      ))}
                      
                      {companies.length === 0 && (
                        <div className="text-center py-12 border-2 border-dashed border-slate-100 rounded-[2rem]">
                          <Building2 className="w-12 h-12 text-slate-200 mx-auto mb-4" />
                          <p className="text-slate-400 font-medium">Aucune société enregistrée.<br/>Importez un fichier Excel pour commencer.</p>
                        </div>
                      )}
                    </div>
                  </section>

                  {/* Selected Company Details (Editable) */}
                  {selectedCompanyId && (
                    <motion.section 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="bg-slate-50 p-8 rounded-[2rem] border border-slate-100 space-y-6"
                    >
                      <h3 className="text-xs font-black text-indigo-600 uppercase tracking-widest">Détails de l'entité sélectionnée</h3>
                      
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div className="space-y-2">
                          <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider ml-1">Nom de l'entreprise</label>
                          <input 
                            type="text" 
                            value={companyInfo.name}
                            onChange={(e) => {
                              const updated = companies.map(c => c.id === selectedCompanyId ? { ...c, name: e.target.value } : c);
                              setCompanies(updated);
                            }}
                            className="w-full px-5 py-3 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-bold text-slate-700"
                          />
                        </div>
                        <div className="space-y-2">
                          <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider ml-1">Email de contact</label>
                          <input 
                            type="email" 
                            value={companyInfo.email}
                            onChange={(e) => {
                              const updated = companies.map(c => c.id === selectedCompanyId ? { ...c, email: e.target.value } : c);
                              setCompanies(updated);
                            }}
                            className="w-full px-5 py-3 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-bold text-slate-700"
                          />
                        </div>
                        <div className="col-span-full space-y-2">
                          <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider ml-1">Adresse complète</label>
                          <textarea 
                            rows={2}
                            value={companyInfo.address}
                            onChange={(e) => {
                              const updated = companies.map(c => c.id === selectedCompanyId ? { ...c, address: e.target.value } : c);
                              setCompanies(updated);
                            }}
                            className="w-full px-5 py-3 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-bold text-slate-700"
                          />
                        </div>
                        <div className="space-y-2">
                          <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider ml-1">N° de TVA</label>
                          <input 
                            type="text" 
                            value={companyInfo.vat}
                            onChange={(e) => {
                              const updated = companies.map(c => c.id === selectedCompanyId ? { ...c, vat: e.target.value } : c);
                              setCompanies(updated);
                            }}
                            className="w-full px-5 py-3 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-bold text-slate-700"
                          />
                        </div>
                        <div className="space-y-2">
                          <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider ml-1">ICE</label>
                          <input 
                            type="text" 
                            value={companyInfo.ice}
                            onChange={(e) => {
                              const updated = companies.map(c => c.id === selectedCompanyId ? { ...c, ice: e.target.value } : c);
                              setCompanies(updated);
                            }}
                            className="w-full px-5 py-3 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-bold text-slate-700"
                          />
                        </div>
                      </div>

                      <div className="pt-4 border-t border-slate-200">
                        <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-wider mb-4">Comptes bancaires enregistrés</h4>
                        <div className="space-y-3">
                          {companyInfo.bankAccount1 && (
                            <div className="p-4 bg-white rounded-xl border border-slate-100 flex items-center justify-between">
                              <div className="flex items-center space-x-3">
                                <Check className="w-4 h-4 text-emerald-500" />
                                <span className="text-xs font-mono text-slate-600">{companyInfo.bankAccount1}</span>
                              </div>
                              <span className="text-[10px] font-black text-slate-300 uppercase">Compte 1</span>
                            </div>
                          )}
                          {companyInfo.bankAccount2 && (
                            <div className="p-4 bg-white rounded-xl border border-slate-100 flex items-center justify-between">
                              <div className="flex items-center space-x-3">
                                <Check className="w-4 h-4 text-emerald-500" />
                                <span className="text-xs font-mono text-slate-600">{companyInfo.bankAccount2}</span>
                              </div>
                              <span className="text-[10px] font-black text-slate-300 uppercase">Compte 2</span>
                            </div>
                          )}
                        </div>
                      </div>
                    </motion.section>
                  )}
                </div>

                <div className="p-8 bg-slate-50 border-t border-slate-100 flex justify-center">
                  <button 
                    onClick={() => setShowSettings(false)}
                    className="px-12 py-4 bg-indigo-600 text-white font-black rounded-2xl hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200"
                  >
                    Fermer
                  </button>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Footer Info */}
        <footer className="mt-12 text-center text-gray-400 text-sm">
          <p>© 2026 FactureGen - Solution de facturation automatisée <span className="ml-2 text-[10px] font-mono bg-slate-100 px-2 py-0.5 rounded-full">v{VERSION}</span></p>
          <div className="mt-4 flex flex-wrap justify-center gap-6">
            <div className="flex items-center space-x-1">
              <div className="w-2 h-2 bg-green-500 rounded-full" />
              <span>Prêt pour Hetzner / PlanetHoster</span>
            </div>
            <div className="flex items-center space-x-1">
              <div className="w-2 h-2 bg-blue-500 rounded-full" />
              <span>Traitement Client-Side Sécurisé</span>
            </div>
            <div className="flex items-center space-x-1">
              <div className="w-2 h-2 bg-purple-500 rounded-full" />
              <span>Matching Intelligent</span>
            </div>
          </div>
        </footer>
      </div>
    </div>
  );
}

// --- Sub-components ---

function FileCard({ title, icon, uploaded, count, onChange, description, unit = "lignes" }: { 
  title: string, 
  icon: React.ReactNode, 
  uploaded: boolean, 
  count: number,
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void,
  description: string,
  unit?: string
}) {
  return (
    <div className={cn(
      "relative bg-white p-8 rounded-[2rem] border transition-all duration-500 card-shadow",
      uploaded ? "border-emerald-100 bg-emerald-50/20" : "border-slate-100 hover:border-indigo-200 hover:shadow-2xl hover:shadow-indigo-50"
    )}>
      <div className="flex items-start justify-between mb-6">
        <div className={cn(
          "p-4 rounded-2xl shadow-sm",
          uploaded ? "bg-emerald-100 text-emerald-600" : "bg-indigo-50 text-indigo-600"
        )}>
          {icon}
        </div>
        {uploaded && (
          <motion.div initial={{ scale: 0 }} animate={{ scale: 1 }} className="flex flex-col items-end">
            <CheckCircle2 className="w-6 h-6 text-emerald-500" />
            <span className="text-[10px] font-black text-emerald-600 mt-1 uppercase tracking-widest">{count} {unit}</span>
          </motion.div>
        )}
      </div>
      <h3 className="font-black text-xl text-slate-800 mb-2">{title}</h3>
      <p className="text-sm text-slate-400 font-medium mb-8 leading-relaxed">{description}</p>
      
      <label className="block">
        <span className="sr-only">Choisir un fichier</span>
        <input 
          type="file" 
          accept=".xlsx, .xls, .csv"
          onChange={onChange}
          className="block w-full text-sm text-slate-500
            file:mr-4 file:py-2.5 file:px-6
            file:rounded-2xl file:border-0
            file:text-sm file:font-bold
            file:bg-indigo-50 file:text-indigo-700
            hover:file:bg-indigo-100
            transition-all cursor-pointer"
        />
      </label>
    </div>
  );
}

function StatusBadge({ active, label }: { active: boolean, label: string }) {
  return (
    <div className={cn(
      "flex items-center space-x-2 px-4 py-2 rounded-2xl border transition-all duration-500",
      active 
        ? "bg-emerald-50 border-emerald-100 text-emerald-600 font-bold" 
        : "bg-slate-50 border-slate-100 text-slate-400 font-medium"
    )}>
      <div className={cn(
        "w-2 h-2 rounded-full",
        active ? "bg-emerald-500 animate-pulse" : "bg-slate-300"
      )} />
      <span className="text-xs uppercase tracking-widest">{label}</span>
    </div>
  );
}
