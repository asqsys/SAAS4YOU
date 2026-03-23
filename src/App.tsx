/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
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
  Users
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---

interface ThirdParty {
  id: string;
  name: string;
  address?: string;
  email?: string;
  vatNumber?: string;
}

interface Service {
  id: string;
  ref: string;
  description: string;
  matchKey: string; // prenom + premiere lettre nom
}

interface BillingLine {
  clientName: string;
  serviceUser: string; // The person name to match with service
  amount: number;
  date: string;
  quantity: number;
  unitPrice: number;
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

  const [companyInfo, setCompanyInfo] = useState({
    name: "",
    address: "",
    email: "",
    vat: ""
  });

  const [showSettings, setShowSettings] = useState(false);

  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [creditDuration, setCreditDuration] = useState(30);

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

  const formatDH = (num: number) => {
    const parts = num.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).split(',');
    const integerPart = parts[0].replace(/\s/g, ' ');
    return `${integerPart},${parts[1]} DH`;
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
        
        setCompanyInfo({
          name,
          address: `${addr}${cp ? ', ' + cp : ''}${city ? ' ' + city : ''}${country ? ', ' + country : ''}`.trim() || "Adresse non renseignée",
          email: String(findValue(row, ['Email', 'Courriel']) || "contact@entreprise.com"),
          vat: String(findValue(row, ['VAT ID', 'TVA', 'Numéro TVA', 'VAT']) || "Non renseigné")
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
            vatNumber: String(findValue(row, ['VAT ID', 'VAT', 'TVA', 'N° TVA']) || '')
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
        const parsed = jsonData.map(row => {
          const id = String(findValue(row, ['ID', 'Ref.', 'Reference', 'Ref', 'Code']) || '');
          const ref = String(findValue(row, ['Ref.', 'Reference', 'Ref', 'Code', 'ID']) || '');
          const description = String(findValue(row, ['Label', 'Description', 'Service', 'Prestation', 'Libellé']) || '');
          return {
            id,
            ref,
            description,
            matchKey: ref.toUpperCase() 
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
        const parsed = jsonData.map(row => {
          const parseNum = (val: any) => {
            if (typeof val === 'number') return val;
            if (typeof val === 'string') {
              const cleaned = val.replace(',', '.').replace(/[^-0-9.]/g, '');
              return cleaned ? Number(cleaned) : 0;
            }
            return 0;
          };

          // Specific columns provided by user
          const unitPrice = parseNum(findValue(row, ['P.U. (excl.)', 'TJM global', 'Unit price (excl.)', 'Prix Unitaire', 'PU', 'Taux', 'TJM']));
          const quantity = parseNum(findValue(row, ['Nbre Jours', 'Nbre de jours', 'Quantite', 'Nb', 'Heures', 'Jours', 'Qté']));
          const amount = parseNum(findValue(row, ['Total HT', 'Factures HT', 'Montant', 'Total', 'HT', 'Net']));
          
          let finalUnitPrice = unitPrice;
          let finalQuantity = quantity;
          let finalAmount = amount;
          
          if (finalAmount === 0 && finalUnitPrice !== 0 && finalQuantity !== 0) {
            finalAmount = finalUnitPrice * finalQuantity;
          }
          if (finalUnitPrice === 0 && finalAmount !== 0 && finalQuantity !== 0) {
            finalUnitPrice = finalAmount / finalQuantity;
          }

          return {
            clientName: String(findValue(row, ['Client', 'Entreprise', 'Société', 'Tiers']) || ''),
            serviceUser: String(findValue(row, ['Description', 'Consultant', 'Collaborateur', 'Nom', 'Prenom Nom', 'Salarié']) || ''),
            amount: finalAmount,
            date: String(findValue(row, ['Date', 'Commentaires', 'Période', 'Mois']) || ''),
            quantity: finalQuantity,
            unitPrice: finalUnitPrice
          };
        }).filter(row => row.serviceUser && row.quantity > 0 && row.unitPrice > 0);
        
        if (jsonData.length === 0) {
          setError(`Le fichier ${type} semble être vide.`);
          return;
        }

        if (parsed.length === 0) {
          setError("Aucune donnée valide trouvée dans le fichier de facturation. Vérifiez que les colonnes 'Unit price (excl.)' et 'Nbre de jours' sont bien remplies et contiennent des valeurs supérieures à 0.");
          return;
        }

        setBillingData(parsed);
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

    setProcessing(true);
    try {
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
      const entries = Object.entries(groupedByClient);
      for (const [clientName, lines] of entries) {
        const clientInfo = thirdPartyData.find(tp => 
          tp.name.toLowerCase().trim() === clientName.toLowerCase().trim() ||
          tp.id.toLowerCase().trim() === clientName.toLowerCase().trim()
        ) || {
          id: 'N/A',
          name: clientName,
          address: 'Adresse non trouvée'
        };

        const invoiceLines = lines.map(line => {
          const nameParts = line.serviceUser.trim().split(/\s+/);
          if (nameParts.length >= 2) {
            const part1 = nameParts[0].toUpperCase();
            const part2 = nameParts[nameParts.length - 1].toUpperCase();
            const key1 = `${part1}-${part2.charAt(0)}`;
            const key2 = `${part2}-${part1.charAt(0)}`;
            
            const foundService = servicesData.find(s => 
              s.ref.toUpperCase().includes(key1) || 
              s.ref.toUpperCase().includes(key2)
            );
            
            if (foundService) {
              return {
                service: foundService,
                quantity: line.quantity,
                unitPrice: line.unitPrice,
                total: line.amount,
                date: line.date
              };
            }
          }

          return {
            service: {
              id: 'ID-MANQUANT',
              ref: 'REF-MANQUANTE',
              description: `Prestation pour ${line.serviceUser}`,
              matchKey: ''
            },
            quantity: line.quantity || 1,
            unitPrice: line.unitPrice || (line.quantity ? line.amount / line.quantity : line.amount),
            total: line.amount,
            date: line.date
          };
        });

        await createPDF(clientInfo, invoiceLines);
      }
      
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

  const createPDF = async (client: any, lines: any[]) => {
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

    const period = `${invoiceDate.getFullYear()}-${String(invoiceDate.getMonth() + 1).padStart(2, '0')}`;

    // Generate Integrity Hash
    const dataToHash = `INV:${invoiceReference}|DATE:${selectedInvoiceDate}|HT:${totalHT.toFixed(2)}|TVA:${tva.toFixed(2)}|TTC:${totalTTC.toFixed(2)}|QTY:${totalQty}`;
    const encoder = new TextEncoder();
    const data = encoder.encode(dataToHash);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    const shortHash = hashHex.substring(0, 16).toUpperCase();

    // Generate QR Code
    const qrContent = `ASQSYS-VERIFY|${dataToHash}|HASH:${hashHex}`;
    const qrDataUrl = await QRCode.toDataURL(qrContent, { margin: 1, width: 100 });

    const drawHeader = (doc: any) => {
      // Color band at the very top
      doc.setFillColor(79, 70, 229); // Indigo 600
      doc.rect(0, 0, 210, 8, 'F');

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

      // Invoice Details - Top Left
      doc.setFontSize(10);
      doc.setTextColor(15, 23, 42);
      doc.setFont("helvetica", "bold");
      doc.text(`Invoice Number: ${invoiceReference}`, 20, 25);
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
      if (client.vatNumber) doc.text(`TVA: ${client.vatNumber}`, 20, 76);
    };

    // Table
    const tableData = lines.map(l => {
      const tva = l.total * 0.20;
      const ttc = l.total + tva;
      return [
        l.service.id,
        `${l.service.description} ${period}`,
        l.quantity,
        formatDH(l.unitPrice),
        formatDH(l.total),
        formatDH(tva),
        formatDH(ttc)
      ];
    });

    autoTable(doc, {
      startY: 80,
      head: [['ID', 'Description', 'Nbre Jours', 'P.U. (excl.)', 'Total HT', 'TVA 20%', 'Montant TTC']],
      body: tableData,
      theme: 'grid',
      headStyles: { 
        fillColor: [79, 70, 229], // Indigo 600
        textColor: 255,
        fontSize: 8,
        fontStyle: 'bold',
        halign: 'center'
      },
      columnStyles: {
        0: { cellWidth: 30 },
        2: { halign: 'center', cellWidth: 15 },
        3: { halign: 'right', cellWidth: 25 },
        4: { halign: 'right', cellWidth: 25 },
        5: { halign: 'right', cellWidth: 25 },
        6: { halign: 'right', cellWidth: 25 }
      },
      styles: { fontSize: 8, cellPadding: 3 },
      margin: { top: 80, bottom: 40 },
      rowPageBreak: 'avoid',
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

    // Totals Box
    const boxWidth = 70;
    const startX = 190 - boxWidth;
    
    doc.setFontSize(10);
    doc.setTextColor(100, 116, 139); // Slate 500
    doc.text(`Total HT:`, startX, finalY + 15);
    doc.text(formatDH(totalHT), 190, finalY + 15, { align: 'right' });
    
    doc.text(`TVA (20%):`, startX, finalY + 22);
    doc.text(formatDH(tva), 190, finalY + 22, { align: 'right' });
    
    doc.setDrawColor(79, 70, 229); // Indigo 600
    doc.setLineWidth(0.5);
    doc.line(startX, finalY + 26, 190, finalY + 26);
    
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.setTextColor(15, 23, 42); // Slate 900
    doc.text(`TOTAL TTC:`, startX, finalY + 33);
    doc.text(formatDH(totalTTC), 190, finalY + 33, { align: 'right' });

    // Amount in words
    doc.setFontSize(9);
    doc.setFont("helvetica", "italic");
    doc.text(`Arrêté la présente facture à la somme de :`, 20, finalY + 45);
    doc.setFont("helvetica", "bolditalic");
    doc.text(`${numberToWordsFR(totalTTC)}`, 20, finalY + 52, { maxWidth: 170 });

    doc.save(`Facture_${client.name.replace(/[^a-z0-9]/gi, '_')}.pdf`);
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
          
          <button 
            onClick={() => setShowSettings(true)}
            className="flex items-center space-x-3 px-8 py-4 bg-white border border-slate-200 rounded-3xl font-bold text-slate-600 hover:bg-slate-50 hover:border-slate-300 transition-all shadow-sm card-shadow"
          >
            <Building2 className="w-6 h-6 text-indigo-600" />
            <span>Ma Société</span>
          </button>
        </header>

        {/* Main Content */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 mb-12">
          <FileCard 
            title="Ma Société" 
            icon={<Building2 className="w-5 h-5" />}
            uploaded={filesUploaded.company}
            count={filesUploaded.company ? 1 : 0}
            onChange={(e) => handleFileUpload(e, 'company')}
            description="Infos émetteur"
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

        {/* Action Section */}
        <div className="bg-white rounded-[2rem] p-10 shadow-sm border border-slate-100 mb-12 card-shadow">
          <div className="flex flex-col items-center justify-center space-y-8">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-8 w-full max-w-4xl">
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
                  Référence
                </label>
                <input 
                  type="text"
                  value={invoiceReference}
                  onChange={(e) => setInvoiceReference(e.target.value)}
                  placeholder="Ex: FAC-2026-001"
                  className="w-full px-5 py-4 bg-slate-50 border border-slate-100 rounded-2xl font-bold text-slate-700 focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all"
                />
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
            </div>

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
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-right">P.U. (excl.)</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-right">Total HT</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-right">TVA 20%</th>
                    <th className="pb-4 font-bold text-slate-400 uppercase tracking-wider text-[10px] text-right">Montant TTC</th>
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
                      isMatched = servicesData.some(s => 
                        s.ref.toUpperCase().includes(`${p1}-${p2.charAt(0)}`) || 
                        s.ref.toUpperCase().includes(`${p2}-${p1.charAt(0)}`)
                      );
                    }
                    return (
                      <tr key={i} className="hover:bg-gray-50/50 transition-colors">
                        <td className="py-4 font-bold text-indigo-600">{line.clientName}</td>
                        <td className="py-4 text-slate-600 font-medium">{line.serviceUser}</td>
                        <td className="py-4 text-right font-mono text-slate-500">{line.quantity}</td>
                        <td className="py-4 text-right font-mono text-slate-500">{formatDH(line.unitPrice)}</td>
                        <td className="py-4 text-right font-bold text-slate-700">{formatDH(line.amount)}</td>
                        <td className="py-4 text-right text-slate-400">{formatDH(line.amount * 0.20)}</td>
                        <td className="py-4 text-right font-bold text-indigo-600">{formatDH(line.amount * 1.20)}</td>
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
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {servicesData.slice(0, 5).map((service, i) => {
                    const invoiceDate = new Date(selectedInvoiceDate);
                    const period = `${invoiceDate.getFullYear()}-${String(invoiceDate.getMonth() + 1).padStart(2, '0')}`;
                    return (
                      <tr key={i} className="hover:bg-gray-50/50 transition-colors">
                        <td className="py-4 font-mono text-blue-600 text-xs">{service.id}</td>
                        <td className="py-4 font-mono text-gray-500 text-xs">{service.ref}</td>
                        <td className="py-4 text-gray-600">{service.description} {period}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </motion.div>
        )}

        {/* Settings Modal */}
        <AnimatePresence>
          {showSettings && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setShowSettings(false)}
                className="absolute inset-0 bg-black/20 backdrop-blur-sm"
              />
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 20 }}
                className="relative bg-white rounded-3xl shadow-2xl w-full max-w-md p-8 overflow-hidden"
              >
                <div className="flex items-center justify-between mb-8">
                  <h2 className="text-2xl font-bold">Ma Société</h2>
                  <button onClick={() => setShowSettings(false)} className="text-gray-400 hover:text-gray-600">
                    <AlertCircle className="w-6 h-6 rotate-45" />
                  </button>
                </div>
                
                <div className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">Nom de l'entreprise</label>
                    <input 
                      type="text" 
                      value={companyInfo.name}
                      onChange={(e) => setCompanyInfo({...companyInfo, name: e.target.value})}
                      className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">Adresse complète</label>
                    <textarea 
                      rows={2}
                      value={companyInfo.address}
                      onChange={(e) => setCompanyInfo({...companyInfo, address: e.target.value})}
                      className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">Email de contact</label>
                    <input 
                      type="email" 
                      value={companyInfo.email}
                      onChange={(e) => setCompanyInfo({...companyInfo, email: e.target.value})}
                      className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">N° de TVA</label>
                    <input 
                      type="text" 
                      value={companyInfo.vat}
                      onChange={(e) => setCompanyInfo({...companyInfo, vat: e.target.value})}
                      className="w-full px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    />
                  </div>
                </div>

                <button 
                  onClick={() => setShowSettings(false)}
                  className="w-full mt-8 py-4 bg-blue-600 text-white font-bold rounded-2xl hover:bg-blue-700 transition-all shadow-lg shadow-blue-100"
                >
                  Enregistrer
                </button>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Footer Info */}
        <footer className="mt-12 text-center text-gray-400 text-sm">
          <p>© 2026 FactureGen - Solution de facturation automatisée</p>
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

function FileCard({ title, icon, uploaded, count, onChange, description }: { 
  title: string, 
  icon: React.ReactNode, 
  uploaded: boolean, 
  count: number,
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void,
  description: string
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
            <span className="text-[10px] font-black text-emerald-600 mt-1 uppercase tracking-widest">{count} lignes</span>
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
