import React, { useState, useEffect } from 'react';
import { Html5QrcodeScanner } from 'html5-qrcode';
import * as XLSX from 'xlsx';

interface StockItem {
  barcode: string;
  labelNo: string;
  labour: number;
  gw: number;
}

interface BillItem extends StockItem {
  rateUsed: number;
  total: number;
}

export default function App() {
  const [activeTab, setActiveTab] = useState<'dashboard' | 'stock' | 'billing'>('dashboard');
  const [goldRate, setGoldRate] = useState<number>(6500);
  const [silverRate, setSilverRate] = useState<number>(85);
  const [inputGold, setInputGold] = useState<string>('6500');
  const [inputSilver, setInputSilver] = useState<string>('85');
  
  const [stockList, setStockList] = useState<StockItem[]>([]);
  const [billItems, setBillItems] = useState<BillItem[]>([]);
  const [barcodeInput, setBarcodeInput] = useState<string>('');
  
  const [custName, setCustName] = useState<string>('');
  const [custPhone, setCustPhone] = useState<string>('');

  useEffect(() => {
    if (activeTab === 'billing') {
      const delayDebounce = setTimeout(() => {
        const scanner = new Html5QrcodeScanner(
          "camera-reader",
          { fps: 15, qrbox: { width: 250, height: 250 } },
          /* verbose= */ false
        );
        
        scanner.render(
          (decodedText) => {
            handleScanSuccess(decodedText);
          },
          (error) => {
            // Scanner error silently handled
          }
        );

        return () => {
          scanner.clear().catch(err => console.error("Failed to clear scanner", err));
        };
      }, 300);

      return () => clearTimeout(delayDebounce);
    }
  }, [activeTab, stockList]);

  const handleRateUpdate = () => {
    setGoldRate(Number(inputGold) || 0);
    setSilverRate(Number(inputSilver) || 0);
    alert("Rates Updated Successfully!");
  };

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const workbook = XLSX.read(bstr, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet) as any[];

      const parsed: StockItem[] = jsonData.map(row => ({
        barcode: String(row['BARCODE'] || '').trim(),
        labelNo: String(row['LABELNO.'] || '').toUpperCase(),
        labour: Number(row['LABOUR'] || 0),
        gw: Number(row['GW.'] || 0)
      })).filter(item => item.barcode !== '');

      setStockList(parsed);
      alert(`🎉 Successfully loaded ${parsed.length} items to active stock!`);
    };
    reader.readAsBinaryString(file);
  };

  const handleScanSuccess = (barcode: string) => {
    const cleanBarcode = barcode.trim();
    const item = stockList.find(s => s.barcode === cleanBarcode);
    
    if (!item) {
      alert("Barcode not found in uploaded stock!");
      return;
    }

    setBillItems(prev => {
      if (prev.some(b => b.barcode === cleanBarcode)) return prev;
      const isSilver = item.labelNo.includes('SILVER') || item.labelNo.includes('CHANDI');
      const rate = isSilver ? silverRate : goldRate;
      const total = (item.gw * rate) + (item.gw * item.labour);
      return [...prev, { ...item, rateUsed: rate, total }];
    });
  };

  const handleManualAdd = () => {
    if (barcodeInput.trim()) {
      handleScanSuccess(barcodeInput.trim());
      setBarcodeInput('');
    }
  };

  const subtotal = billItems.reduce((acc, item) => acc + item.total, 0);
  const gst = subtotal * 0.03;
  const grandTotal = subtotal + gst;

  return (
    <div className="flex h-screen overflow-hidden bg-slate-50 font-sans antialiased">
      {/* SIDEBAR */}
      <div className="w-64 bg-slate-900 text-white flex flex-col shadow-xl print:hidden">
        <div className="p-5 border-b border-slate-800 bg-slate-950">
          <h2 className="text-xl font-black text-amber-400 tracking-wider">SALT & GLITZ</h2>
          <p className="text-[10px] text-slate-400 mt-0.5 tracking-widest uppercase">Premium Jewellery Platform</p>
        </div>
        <nav className="flex-1 p-4 space-y-2">
          <button 
            onClick={() => setActiveTab('dashboard')} 
            className={`w-full text-left py-2.5 px-4 rounded-lg text-sm font-medium transition ${activeTab === 'dashboard' ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-300 hover:bg-slate-800'}`}>
            🏪 Store Dashboard
          </button>
          <button 
            onClick={() => setActiveTab('stock')} 
            className={`w-full text-left py-2.5 px-4 rounded-lg text-sm font-medium transition ${activeTab === 'stock' ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-300 hover:bg-slate-800'}`}>
            📦 Stock Matrix Import
          </button>
          <button 
            onClick={() => setActiveTab('billing')} 
            className={`w-full text-left py-2.5 px-4 rounded-lg text-sm font-medium transition ${activeTab === 'billing' ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-300 hover:bg-slate-800'}`}>
            🛍️ Invoice / Estimate
          </button>
        </nav>
        <div className="p-4 border-t border-slate-800 text-center text-xs text-slate-500">FY: 2026-27</div>
      </div>

      {/* MAIN PANEL */}
      <div className="flex-1 overflow-y-auto p-8 print:p-0">
        
        {/* DASHBOARD TAB */}
        {activeTab === 'dashboard' && (
          <div className="max-w-4xl mx-auto space-y-6">
            <div className="bg-white p-6 rounded-xl shadow-sm border">
              <h2 className="text-xl font-bold text-gray-800">Live Precious Metal Rates Today</h2>
              <div className="grid grid-cols-2 gap-4 mt-4">
                <div className="bg-gradient-to-r from-amber-500 to-yellow-600 p-6 rounded-xl text-white font-bold text-lg shadow-sm">
                  Gold (24K): ₹{goldRate.toLocaleString()} / g
                </div>
                <div className="bg-gradient-to-r from-slate-400 to-slate-600 p-6 rounded-xl text-white font-bold text-lg shadow-sm">
                  Silver: ₹{silverRate.toLocaleString()} / g
                </div>
              </div>
            </div>
            <div className="bg-white p-6 rounded-xl shadow-sm border max-w-md">
              <h3 className="font-bold text-gray-700 mb-4">Update Live Market Rates</h3>
              <input 
                type="number" 
                value={inputGold} 
                onChange={(e) => setInputGold(e.target.value)}
                className="w-full border px-3 py-2 rounded-lg text-sm mb-3 outline-none focus:border-indigo-500" 
                placeholder="Gold Rate" 
              />
              <input 
                type="number" 
                value={inputSilver} 
                onChange={(e) => setInputSilver(e.target.value)}
                className="w-full border px-3 py-2 rounded-lg text-sm mb-3 outline-none focus:border-indigo-500" 
                placeholder="Silver Rate" 
              />
              <button 
                onClick={handleRateUpdate}
                className="w-full bg-indigo-600 text-white py-2 rounded-lg text-sm font-bold shadow-sm hover:bg-indigo-700 transition">
                Save Rates
              </button>
            </div>
          </div>
        )}

        {/* STOCK TAB */}
        {activeTab === 'stock' && (
          <div className="max-w-4xl mx-auto bg-white p-6 rounded-xl shadow-sm border">
            <div className="flex justify-between items-center border-b pb-4 mb-4">
              <div>
                <h2 className="text-lg font-bold text-gray-800">Excel Stock Ledger Sync</h2>
                <p className="text-xs text-gray-400">Headers required: BARCODE, LABELNO., LABOUR, GW.</p>
              </div>
              <input 
                type="file" 
                accept=".xlsx, .xls" 
                onChange={handleExcelUpload}
                className="text-sm cursor-pointer" 
              />
            </div>
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-slate-50 text-slate-600 font-bold uppercase">
                <tr>
                  <th className="p-3">Barcode Serial</th>
                  <th className="p-3">Label</th>
                  <th className="p-3 text-right">Gross Wt.</th>
                  <th className="p-3 text-right">Labour / g</th>
                </tr>
              </thead>
              <tbody className="divide-y text-slate-700">
                {stockList.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="p-8 text-center text-gray-400">
                      No active stock items registered. Please select and upload excel sheet.
                    </td>
                  </tr>
                ) : (
                  stockList.map((item, idx) => (
                    <tr key={idx} className="hover:bg-slate-50">
                      <td className="p-3 font-mono font-bold">{item.barcode}</td>
                      <td className="p-3">{item.labelNo}</td>
                      <td className="p-3 text-right">{item.gw.toFixed(3)} g</td>
                      <td className="p-3 text-right">₹{item.labour}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* BILLING TAB */}
        {activeTab === 'billing' && (
          <div className="max-w-3xl mx-auto bg-white p-8 rounded-xl shadow-md border print:border-none print:shadow-none print:p-0">
            <div className="flex justify-between items-start border-b-2 border-slate-900 pb-4 mb-4">
              <div>
                <h2 className="text-2xl font-black text-slate-800 tracking-tight">TAX INVOICE / ESTIMATE</h2>
                <p className="text-xs font-mono text-gray-500 font-bold mt-1">Invoice Series: SALT/2026-27/001</p>
              </div>
              <div className="text-right">
                <h2 className="text-xl font-black text-indigo-900 tracking-wide">SALT & GLITZ</h2>
                <p className="text-[11px] text-gray-500 font-medium">GSTIN: 24AAAAA0000A1Z1 | Surat</p>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-4 mb-4 bg-slate-50 p-3 rounded-lg border print:hidden">
              <input 
                type="text" 
                placeholder="Client Full Name" 
                value={custName}
                onChange={(e) => setCustName(e.target.value)}
                className="bg-white border p-2 rounded text-sm outline-none focus:border-indigo-500" 
              />
              <input 
                type="text" 
                placeholder="Mobile Number" 
                value={custPhone}
                onChange={(e) => setCustPhone(e.target.value)}
                className="bg-white border p-2 rounded text-sm outline-none focus:border-indigo-500" 
              />
            </div>

            {/* 📸 LIVE SCANNER WINDOW */}
            <div className="mb-4 print:hidden bg-slate-100 p-4 rounded-xl border border-dashed border-slate-300 max-w-xs mx-auto text-center">
              <p className="text-[10px] font-black text-indigo-600 mb-2">📸 Live Laptop/Mobile Web Scanner</p>
              <div id="camera-reader" className="w-full overflow-hidden rounded-md bg-black"></div>
            </div>

            <div className="flex gap-2 mb-4 print:hidden">
              <input 
                type="text" 
                placeholder="Type barcode and hit Enter..." 
                value={barcodeInput}
                onChange={(e) => setBarcodeInput(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleManualAdd()}
                className="flex-1 border px-3 py-2 rounded-lg text-sm outline-none focus:border-indigo-500" 
              />
              <button onClick={handleManualAdd} className="bg-slate-800 text-white px-5 rounded-lg text-sm font-semibold hover:bg-slate-700 transition">Add</button>
              <button onClick={() => window.print()} className="bg-indigo-600 text-white px-5 rounded-lg text-sm font-bold shadow-md hover:bg-indigo-700 transition">🖨️ Print / Save PDF</button>
            </div>

            {/* BILL CLIENT DATA FOR PRINT */}
            {(custName || custPhone) && (
              <div className="hidden print:block mb-4 text-xs bg-slate-50 p-2 border rounded">
                {custName && <div><strong>Customer:</strong> {custName}</div>}
                {custPhone && <div><strong>Mobile:</strong> {custPhone}</div>}
              </div>
            )}

            <table className="w-full text-left text-xs border border-collapse border-slate-200">
              <thead className="bg-slate-900 text-white font-bold uppercase text-[10px]">
                <tr>
                  <th className="p-2 border">Item Details</th>
                  <th className="p-2 border text-right">Gross Wt.</th>
                  <th className="p-2 border text-right">Rate Used</th>
                  <th className="p-2 border text-right">Labour/g</th>
                  <th className="p-2 border text-right">Total</th>
                </tr>
              </thead>
              <tbody className="divide-y text-slate-800 font-medium">
                {billItems.length === 0 ? (
                  <tr>
                    <td colSpan={5} className="p-6 text-center text-gray-400">
                      No items loaded into active bill matrix. Use barcode scanner or manual input.
                    </td>
                  </tr>
                ) : (
                  billItems.map((item, idx) => (
                    <tr key={idx}>
                      <td className="p-2 border font-mono font-bold">
                        {item.barcode}
                        <span className="text-[10px] text-slate-400 block font-sans font-normal">({item.labelNo})</span>
                      </td>
                      <td className="p-2 border text-right">{item.gw.toFixed(3)} g</td>
                      <td className="p-2 border text-right">₹{item.rateUsed}</td>
                      <td className="p-2 border text-right">₹{item.labour}</td>
                      <td className="p-2 border text-right font-mono font-bold">₹{Math.round(item.total).toLocaleString()}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>

            <div className="flex justify-between items-start pt-4 mt-2 border-t border-slate-300">
              <div className="text-[10px] text-slate-400 max-w-xs space-y-1">
                <p className="font-bold text-slate-600">Terms & Regulatory Clauses:</p>
                <p>1. Custom manufactured goods cannot be exchanged or returned.</p>
                <p>2. Subject to Surat Jurisdiction. Computer Generated Document. E&OE.</p>
              </div>
              <div className="w-48 text-xs text-slate-600 font-bold space-y-1 text-right">
                <div>Subtotal Gross: ₹{Math.round(subtotal).toLocaleString()}</div>
                <div>Total GST (3%): ₹{Math.round(gst).toLocaleString()}</div>
                <div className="text-sm font-black text-indigo-900 border-t pt-1">
                  Grand Total: ₹{Math.round(grandTotal).toLocaleString()}
                </div>
              </div>
            </div>

            <div className="hidden print:flex justify-between mt-12 text-[10px]">
              <div className="border-t border-slate-300 border-dashed w-36 text-center pt-1 text-slate-400">Customer Signature</div>
              <div className="border-t border-slate-300 border-dashed w-52 text-center pt-1 font-bold text-slate-900">
                For, SALT & GLITZ<br /><br /><span className="text-[9px] font-normal text-slate-400">Authorized Signatory</span>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
