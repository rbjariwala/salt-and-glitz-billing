import React, { useState, useEffect, ChangeEvent } from 'react';
import * as XLSX from 'xlsx';
// કેમેરા સ્કેનર માટેની ઓફિશિયલ લાઈબ્રેરી
import { Html5QrcodeScanner } from "html5-qrcode";

// Data Interfaces
interface StockItem { id: string; barcode: string; labelNo: string; labour: number; gw: number; status: 'Available' | 'Sold'; itemType: 'Gold' | 'Silver'; }
interface InvoiceItem { barcode: string; labelNo: string; gw: number; labour: number; rateUsed: number; total: number; hsn: string; }
interface HistoryItem { id: string; date: string; type: 'Sale' | 'Estimate'; amount: number; itemsCount: number; status: string; }
interface CustomerProfile { name: string; mobile: string; totalPurchase: number; history: HistoryItem[]; }

export default function App() {
  const [currentTab, setCurrentTab] = useState<'dashboard' | 'stock' | 'estimate' | 'sale' | 'ledger' | 'reports'>('dashboard');
  
  // Live Metal Rates (With LocalStorage backup)
  const [goldRate, setGoldRate] = useState<number>(() => Number(localStorage.getItem('goldRate')) || 6500);
  const [silverRate, setSilverRate] = useState<number>(() => Number(localStorage.getItem('silverRate')) || 85);
  const [inputGold, setInputGold] = useState(goldRate);
  const [inputSilver, setInputSilver] = useState(silverRate);

  // Shop Profile Configuration (SALT & GLITZ કાયમ માટે ફિક્સ સેટ)
  const [shopName] = useState('SALT & GLITZ');
  const [shopPhone, setShopPhone] = useState(() => localStorage.getItem('shopPhone') || '9876543210');
  const [shopAddress, setShopAddress] = useState(() => localStorage.getItem('shopAddress') || 'Surat, Gujarat');
  const [shopGST, setShopGST] = useState(() => localStorage.getItem('shopGST') || '24AAAAA0000A1Z1');
  const [financialYear, setFinancialYear] = useState(() => localStorage.getItem('financialYear') || '2026-27');

  const [inputShopPhone, setInputShopPhone] = useState(shopPhone);
  const [inputShopAddress, setInputShopAddress] = useState(shopAddress);
  const [inputShopGST, setInputShopGST] = useState(shopGST);

  // 📸 શોપ લોગો સ્ટેટ (કમ્પ્યુટરની મેમરીમાં સ્ટોર થશે)
  const [shopLogo, setShopLogo] = useState<string | null>(() => localStorage.getItem('shopLogo') || null);

  // Master Stock & Billing States
  const [stockList, setStockList] = useState<StockItem[]>([]);
  const [searchBarcode, setSearchBarcode] = useState('');
  const [searchLabel, setSearchLabel] = useState('');
  const [customerName, setCustomerName] = useState('');
  const [mobileNumber, setMobileNumber] = useState('');
  const [billingBarcode, setBillingBarcode] = useState('');
  const [billItems, setBillItems] = useState<InvoiceItem[]>([]);
  const [discount, setDiscount] = useState<number>(0);
  const [invoiceNumber, setInvoiceNumber] = useState('');
  const [invoiceReports, setInvoiceReports] = useState<any[]>([]);

  // 📸 ઓટોમેશન: લેપટોપ કેમેરા સ્કેનર એક્ટિવેશન[span_1](start_span)[span_1](end_span)
  useEffect(() => {
    if (currentTab === 'sale' || currentTab === 'estimate') {
      const scanner = new Html5QrcodeScanner("camera-reader", { fps: 10, qrbox: 250 }, false);
      scanner.render((decodedText) => {
        handleAutoScanning(decodedText, currentTab === 'estimate');
      }, (error) => { /* Frame Skip Fallback */ });
      
      return () => { scanner.clear().catch(err => console.error(err)); };
    }
  }, [currentTab, stockList, goldRate, silverRate]);

  useEffect(() => {
    const prefix = currentTab === 'sale' ? 'INV' : 'EST';
    setInvoiceNumber(`${prefix}-${financialYear}-${Math.floor(10000 + Math.random() * 90000)}`);
  }, [currentTab, financialYear]);

  // લોગો અપલોડ હેન્ડલર
  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64String = reader.result as string;
        setShopLogo(base64String);
        localStorage.setItem('shopLogo', base64String);
      };
      reader.readAsDataURL(file);
    }
  };

  // Shop Information Updater
  const saveShopDetails = (e: React.FormEvent) => {
    e.preventDefault();
    setShopPhone(inputShopPhone); setShopAddress(inputShopAddress); setShopGST(inputShopGST);
    localStorage.setItem('shopPhone', inputShopPhone);
    localStorage.setItem('shopAddress', inputShopAddress); localStorage.setItem('shopGST', inputShopGST);
    alert("Store Profile updated successfully!");
  };

  // Metal Rate Updater
  const saveRates = (e: React.FormEvent) => {
    e.preventDefault();
    setGoldRate(inputGold); setSilverRate(inputSilver);
    localStorage.setItem('goldRate', inputGold.toString()); localStorage.setItem('silverRate', inputSilver.toString());
    alert("Live Metal Rates saved successfully!");
  };

  // Financial Year Closing Tool
  const handleFinancialYearChange = () => {
    const currentYears = financialYear.split('-');
    const nextFY = `${parseInt(currentYears[0]) + 1}-${(parseInt(currentYears[1]) + 1).toString().slice(-2)}`;
    if (confirm(`⚠️ Close year and shift to "${nextFY}"? Available stock will carry forward.`)) {
      setFinancialYear(nextFY); localStorage.setItem('financialYear', nextFY);
      setStockList(stockList.filter(item => item.status === 'Available'));
      setInvoiceReports([]); alert(`Financial Year shifted to ${nextFY}!`);
    }
  };

  // Excel Stock Import Matrix
  const handleExcelImport = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const data = XLSX.utils.sheet_to_json<any>(wb.Sheets[wb.SheetNames[0]]);

      const imported: StockItem[] = data.map((row, idx) => {
        const label = String(row['LABELNO.'] || '').toUpperCase();
        return {
          id: `STK-${Date.now()}-${idx}`,
          barcode: String(row['BARCODE'] || '').trim(),
          labelNo: label,
          labour: Number(row['LABOUR'] || 0),
          gw: Number(row['GW.'] || 0),
          status: 'Available',
          itemType: label.includes('SILVER') || label.includes('CHANDI') ? 'Silver' : 'Gold'
        };
      }).filter(i => i.barcode !== '');

      setStockList([...stockList, ...imported]);
      alert(`🎉 Successfully imported ${imported.length} items to stock!`);
    };
    reader.readAsBinaryString(file);
  };

  // 🛠️ ગણતરી પદ્ધતિ: કેમેરા સ્કેન અથવા હાથેથી બારકોડ એન્ટ્રી સબમિશન[span_2](start_span)[span_2](end_span)
  const handleAutoScanning = (scannedCode: string, isEstimate: boolean) => {
    const item = stockList.find(s => s.barcode === scannedCode && s.status === 'Available');
    if (!item) return;

    const isSilver = item.itemType === 'Silver';
    const currentRate = isSilver ? silverRate : goldRate;
    const defaultHSN = isSilver ? '7114' : '7113';

    // પાકી મજૂરી ગણતરી: (વજન * રેટ) + (વજન * બારકોડવાળી મજૂરી)[span_3](start_span)[span_3](end_span)
    const lineTotal = (item.gw * currentRate) + (item.gw * item.labour);

    const newItem: InvoiceItem = {
      barcode: item.barcode, labelNo: item.labelNo, gw: item.gw,
      labour: item.labour, rateUsed: currentRate, total: lineTotal, hsn: defaultHSN
    };

    setBillItems(prev => {
      if (prev.some(b => b.barcode === scannedCode)) return prev;
      if (!isEstimate) {
        setStockList(prevStock => prevStock.map(s => s.barcode === item.barcode ? { ...s, status: 'Sold' } : s));
      }
      return [...prev, newItem];
    });
  };

  const subTotal = billItems.reduce((sum, item) => sum + item.total, 0);
  const afterDiscount = Math.max(0, subTotal - discount);
  const gst = currentTab === 'sale' ? afterDiscount * 0.03 : 0;
  const grandTotal = afterDiscount + gst;

  return (
    <div className="flex h-screen bg-slate-50 font-sans print:bg-white overflow-hidden">
      
      {/* SIDEBAR NAVIGATION[span_4](start_span)[span_4](end_span) */}
      <div className="w-64 bg-slate-900 text-white flex flex-col print:hidden shadow-xl">
        <div className="p-5 border-b border-slate-800 bg-slate-950">
          <h2 className="text-lg font-black text-amber-400 tracking-wider">SALT & GLITZ</h2>
          <p className="text-xs text-slate-400 mt-0.5">Premium Billing v1.0</p>
        </div>
        <nav className="flex-1 p-4 space-y-1.5">
          {[
            { id: 'dashboard', label: '🏪 Dashboard' },
            { id: 'stock', label: '📦 Stock Import' },
            { id: 'estimate', label: '📝 Estimate / Quotation' },
            { id: 'sale', label: '🛍️ Tax Invoice (Sale)' }
          ].map(tab => (
            <button key={tab.id} onClick={() => { setCurrentTab(tab.id as any); setBillItems([]); setDiscount(0); }} className={`w-full text-left py-2.5 px-4 rounded-lg text-sm font-medium transition ${currentTab === tab.id ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-300 hover:bg-slate-800'}`}>{tab.label}</button>
          ))}
        </nav>
      </div>

      {/* MAIN VIEW */}
      <div className="flex-1 overflow-y-auto print:p-0">
        
        {/* DASHBOARD TAB */}
        {currentTab === 'dashboard' && (
          <div className="p-8 max-w-5xl mx-auto space-y-6">
            <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
              <h2 className="text-xl font-bold text-gray-800">Current Metal Rates Today</h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mt-4">
                <div className="bg-gradient-to-br from-amber-500 to-yellow-600 p-6 rounded-xl text-white shadow-md">
                  <p className="text-xs uppercase font-bold text-amber-100">Gold Rate</p>
                  <p className="text-4xl font-black mt-2">₹ {goldRate.toLocaleString()} <span className="text-xs font-normal">/ per Gram</span></p>
                </div>
                <div className="bg-gradient-to-br from-slate-400 to-slate-600 p-6 rounded-xl text-white shadow-md">
                  <p className="text-xs uppercase font-bold text-slate-100">Silver Rate</p>
                  <p className="text-4xl font-black mt-2">₹ {silverRate.toLocaleString()} <span className="text-xs font-normal">/ per Gram</span></p>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
                <h3 className="font-bold text-gray-700 text-sm mb-4">Update Metal Rates</h3>
                <form onSubmit={saveRates} className="space-y-4">
                  <input type="number" placeholder="Gold Rate" value={inputGold} onChange={e => setInputGold(Number(e.target.value))} className="w-full border px-3 py-2 rounded-lg text-sm outline-none" />
                  <input type="number" placeholder="Silver Rate" value={inputSilver} onChange={e => setInputSilver(Number(e.target.value))} className="w-full border px-3 py-2 rounded-lg text-sm outline-none" />
                  <button type="submit" className="w-full bg-indigo-600 text-white py-2 rounded-lg text-sm font-semibold hover:bg-indigo-700">Save Rates</button>
                </form>
              </div>

              <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
                <h3 className="font-bold text-gray-700 text-sm mb-4">Store Settings & Logo</h3>
                <div className="mb-4 bg-slate-50 p-3 rounded-lg border">
                  <label className="block text-xs font-bold text-gray-600 mb-1">Upload Shop Logo:</label>
                  <input type="file" accept="image/*" onChange={handleLogoUpload} className="text-xs" />
                  {shopLogo && <p className="text-xs text-green-600 font-bold mt-1">✓ Logo loaded inside database!</p>}
                </div>
                <form onSubmit={saveShopDetails} className="space-y-3">
                  <input type="text" placeholder="Phone Number" value={inputShopPhone} onChange={e => setInputShopPhone(e.target.value)} className="w-full border px-3 py-1.5 rounded-lg text-sm outline-none" />
                  <input type="text" placeholder="GST Number" value={inputShopGST} onChange={e => setInputShopGST(e.target.value)} className="w-full border px-3 py-1.5 rounded-lg text-sm outline-none" />
                  <textarea placeholder="Address" value={inputShopAddress} onChange={e => setInputShopAddress(e.target.value)} className="w-full border px-3 py-1.5 rounded-lg text-sm outline-none h-12" />
                  <button type="submit" className="w-full bg-emerald-600 text-white py-2 rounded-lg text-sm font-semibold hover:bg-emerald-700">Update Profile</button>
                </form>
              </div>
            </div>

            <div className="bg-white p-6 rounded-xl shadow-sm border border-red-100 max-w-sm">
              <h3 className="font-bold text-gray-800 text-sm mb-1">Financial Year: <span className="text-indigo-600">{financialYear}</span></h3>
              <button onClick={handleFinancialYearChange} className="w-full bg-rose-600 text-white py-2 rounded-lg text-xs font-bold mt-2 hover:bg-rose-700">Close Financial Year</button>
            </div>
          </div>
        )}

        {/* STOCK IMPORT TAB */}
        {currentTab === 'stock' && (
          <div className="p-8 max-w-6xl mx-auto space-y-6">
            <div className="flex justify-between items-center bg-white p-6 rounded-xl shadow-sm border">
              <div>
                <h2 className="text-lg font-bold text-gray-800">Excel Stock Import Matrix</h2>
                <p className="text-xs text-gray-400 mt-1">Columns: BARCODE, LABELNO., LABOUR, GW.</p>
              </div>
              <input type="file" accept=".xlsx, .xls" onChange={handleExcelImport} className="text-sm cursor-pointer" />
            </div>

            <div className="bg-white rounded-xl shadow-sm overflow-hidden border">
              <table className="w-full text-left border-collapse">
                <thead className="bg-gray-50 text-xs font-bold text-gray-500 uppercase border-b">
                  <tr>
                    <th className="p-4">Barcode</th>
                    <th className="p-4">Label Details</th>
                    <th className="p-4">Metal Type</th>
                    <th className="p-4 text-right">Weight (GW)</th>
                    <th className="p-4 text-right">Labour / g</th>
                    <th className="p-4 text-center">Status</th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y">
                  {stockList.map((item, idx) => (
                    <tr key={idx}>
                      <td className="p-4 font-mono font-bold text-slate-900">{item.barcode}</td>
                      <td className="p-4 text-gray-600">{item.labelNo}</td>
                      <td className="p-4"><span className={`text-xs font-bold px-2 py-0.5 rounded ${item.itemType === 'Silver' ? 'bg-slate-100 text-slate-700' : 'bg-amber-100 text-amber-800'}`}>{item.itemType}</span></td>
                      <td className="p-4 text-right font-medium">{item.gw.toFixed(3)} g</td>
                      <td className="p-4 text-right">₹{item.labour}</td>
                      <td className="p-4 text-center">
                        <span className={`px-2.5 py-0.5 text-xs font-bold rounded-full ${item.status === 'Available' ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>{item.status}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* BILLING & ESTIMATE BILLS TAB[span_5](start_span)[span_5](end_span) */}
        {(currentTab === 'estimate' || currentTab === 'sale') && (
          <div className="p-8 max-w-4xl mx-auto print:p-0">
            <div className="bg-white p-8 rounded-xl shadow-md border print:shadow-none print:border-none">
              
              <div className="flex justify-between items-start border-b pb-6 mb-6">
                <div>
                  <h2 className="text-2xl font-black text-slate-800 tracking-wide">{currentTab === 'sale' ? 'TAX INVOICE' : 'ESTIMATE QUOTATION'}</h2>
                  <p className="text-sm font-semibold text-gray-600 mt-1">Invoice Series: <span className="font-mono text-gray-900">{invoiceNumber}</span></p>
                </div>
                <div className="text-right">
                  {/* જો લોગો અપલોડ કર્યો હશે તો લોગો દેખાશે, નહીંતર ફિક્સ નામ છપાશે */}
                  {shopLogo ? (
                    <img src={shopLogo} alt="Logo" className="h-14 w-auto ml-auto mb-1 object-contain" />
                  ) : (
                    <h2 className="text-xl font-black text-slate-900">{shopName}</h2>
                  )}
                  <p className="text-xs text-gray-500">GSTIN: {shopGST} | Ph: {shopPhone}</p>
                  <p className="text-xs text-gray-500">{shopAddress}</p>
                </div>
              </div>

              {/* Customer Info */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6 print:mb-4 bg-slate-50 p-4 rounded-lg print:bg-transparent print:p-0">
                <input type="text" placeholder="Customer Name" value={customerName} onChange={e => setCustomerName(e.target.value)} className="border px-3 py-1.5 rounded text-sm outline-none print:border-none print:p-0 print:font-bold" />
                <input type="text" placeholder="Mobile Number" value={mobileNumber} onChange={e => setMobileNumber(e.target.value)} className="border px-3 py-1.5 rounded text-sm outline-none print:border-none print:p-0" />
              </div>

              {/* 📸 લેપટોપ કેમેરા સ્કેનર વિન્ડો[span_6](start_span)[span_6](end_span) */}
              <div className="mb-6 print:hidden bg-slate-100 p-4 rounded-xl border border-dashed border-gray-300 max-w-sm mx-auto">
                <p className="text-xs text-center font-bold text-indigo-600 mb-2">📸 Live Laptop Cam Scanner Active</p>
                <div id="camera-reader" className="w-full overflow-hidden rounded-lg"></div>
              </div>

              {/* મેન્યુઅલ બેકઅપ સબમિટ */}
              <div className="flex justify-between items-center mb-4 print:hidden">
                <form onSubmit={(e) => { e.preventDefault(); handleAutoScanning(billingBarcode, currentTab === 'estimate'); setBillingBarcode(''); }} className="flex-1 flex gap-2">
                  <input type="text" placeholder="Type Barcode if camera is not focusing..." value={billingBarcode} onChange={e => setBillingBarcode(e.target.value)} className="flex-1 border px-3 py-2 rounded-lg text-sm outline-none" />
                  <button type="submit" className="bg-slate-800 text-white px-5 rounded-lg text-sm font-medium">Add</button>
                </form>
                <button onClick={() => window.print()} className="ml-4 bg-indigo-600 text-white px-5 py-2 rounded-lg text-sm font-bold shadow-sm">🖨️ Print / Save PDF</button>
              </div>

              {/* Line Items View[span_7](start_span)[span_7](end_span) */}
              <table className="w-full text-left border-collapse border rounded-lg overflow-hidden mb-6">
                <thead className="bg-slate-50 border-b text-xs uppercase font-bold text-gray-500">
                  <tr>
                    <th className="p-3">Item Details</th>
                    <th className="p-3 text-center">HSN</th>
                    <th className="p-3 text-right">Weight (GW)</th>
                    <th className="p-3 text-right">Rate Used</th>
                    <th className="p-3 text-right">Labour / g</th>
                    <th className="p-3 text-right">Amount</th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y">
                  {billItems.map((item, i) => (
                    <tr key={i}>
                      <td className="p-3 font-mono font-medium text-slate-900">{item.barcode} <span className="text-xs text-gray-400 block">({item.labelNo})</span></td>
                      <td className="p-3 text-center font-bold text-gray-600">{item.hsn}</td>
                      <td className="p-3 text-right font-medium">{item.gw.toFixed(3)} g</td>
                      <td className="p-3 text-right">₹{item.rateUsed}</td>
                      <td className="p-3 text-right">₹{item.labour}</td>
                      <td className="p-3 text-right font-bold text-slate-900">₹{item.total.toLocaleString('en-IN')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Financial Breakdowns[span_8](start_span)[span_8](end_span) */}
              <div className="flex justify-between items-start pt-4 border-t">
                <div className="text-xs text-gray-400 max-w-sm">
                  <p className="font-bold text-gray-700 uppercase mb-1">Terms & Conditions:</p>
                  <p>1. Goods once sold will not be taken back or exchanged.</p>
                  <p>2. Subject to Surat Jurisdiction.</p>
                  <p>3. This is a computer-generated tax invoice. E&OE.</p>
                </div>
                <div className="w-64 space-y-2 text-sm text-gray-600">
                  <div className="flex justify-between"><span>Subtotal:</span><span className="font-medium text-slate-900">₹{subTotal.toLocaleString('en-IN')}</span></div>
                  <div className="flex justify-between text-rose-600 items-center">
                    <span>Discount (₹):</span>
                    <input type="number" value={discount || ''} onChange={e => setDiscount(Number(e.target.value))} className="w-20 border text-right px-1 print:hidden" />
                    <span className="hidden print:block">- ₹{discount}</span>
                  </div>
                  {currentTab === 'sale' && (
                    <>
                      <div className="flex justify-between text-xs text-gray-400"><span>CGST (1.5%):</span><span>₹{(gst/2).toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between text-xs text-gray-400"><span>SGST (1.5%):</span><span>₹{(gst/2).toLocaleString('en-IN')}</span></div>
                      <div className="flex justify-between font-medium"><span>Total GST (3%):</span><span className="text-slate-900">₹{gst.toLocaleString('en-IN')}</span></div>
                    </>
                  )}
                  <div className="flex justify-between text-base font-black text-indigo-900 border-t pt-2"><span>Grand Total:</span><span>₹{Math.round(grandTotal).toLocaleString('en-IN')}</span></div>
                </div>
              </div>

              {/* Signatures[span_9](start_span)[span_9](end_span) */}
              <div className="hidden print:flex justify-between mt-16 text-xs pt-4">
                <div className="border-t dashed w-40 text-center pt-1 text-gray-500">Customer Signature</div>
                <div className="border-t dashed w-56 text-center pt-1 font-bold text-slate-900">For, SALT & GLITZ<br><br><span className="text-[10px] font-normal text-gray-400">Authorized Signatory</span></div>
              </div>

            </div>
          </div>
        )}

      </div>
    </div>
  );
}
