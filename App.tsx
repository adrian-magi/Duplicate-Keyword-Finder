import React, { useState, useCallback, useMemo } from 'react';
import { useDropzone, DropzoneOptions } from 'react-dropzone';
import { 
  CheckCircle2, 
  XCircle, 
  Loader2, 
  Upload, 
  AlertCircle,
  RefreshCw,
  FileSpreadsheet,
  Search,
  Target
} from 'lucide-react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface PpcDuplicateGroup {
  keyword: string;
  matchType: string;
  asin: string;
  placement: string;
  audience: string;
  instances: {
    campaign: string;
    adGroup: string;
    campaignState: string;
    adGroupState: string;
    acos: string;
    bid: string;
    impressions: string;
    clicks: string;
    spend: string;
    sales: string;
    orders: string;
    roas: string;
    row: number;
  }[];
}

export default function App() {
  // PPC Deduplicator State
  const [ppcFileName, setPpcFileName] = useState<string | null>(null);
  const [isAuditingPpc, setIsAuditingPpc] = useState(false);
  const [ppcAuditResults, setPpcAuditResults] = useState<PpcDuplicateGroup[] | null>(null);
  const [ppcError, setPpcError] = useState<string | null>(null);

  const resetPpcAudit = () => {
    setPpcFileName(null);
    setPpcAuditResults(null);
    setPpcError(null);
  };  const processPpcData = (data: any[], isArrayRows: boolean = false) => {
    try {
      const groups: Record<string, PpcDuplicateGroup> = {};
      const campaignStates: Record<string, string> = {};
      const adGroupStates: Record<string, string> = {};
      const adGroupToCampaign: Record<string, string> = {};
      const campaignPlacementProfiles: Record<string, Record<string, string>> = {};
      const adGroupAsins: Record<string, Set<string>> = {};
      const adGroupAudiences: Record<string, string> = {};

      let headers: string[] = [];
      let headerMap: Record<string, number> = {};

      if (isArrayRows && data.length > 0) {
        headers = data[0].map((h: any) => String(h || '').toLowerCase());
        headers.forEach((h, i) => {
          if (h) headerMap[h] = i;
        });
      }

      const getVal = (row: any, searchTerms: string[]) => {
        if (isArrayRows) {
          for (const term of searchTerms) {
            const t = term.toLowerCase();
            const idx = headerMap[t];
            if (idx !== undefined && row[idx] !== undefined && row[idx] !== null && String(row[idx]).trim() !== '') return row[idx];
          }
          // Fallback to includes
          for (const term of searchTerms) {
            const t = term.toLowerCase();
            const matchIdx = headers.findIndex(h => h.includes(t));
            if (matchIdx !== -1 && row[matchIdx] !== undefined && row[matchIdx] !== null && String(row[matchIdx]).trim() !== '') return row[matchIdx];
          }
          return null;
        } else {
          const keys = Object.keys(row);
          for (const term of searchTerms) {
            const t = term.toLowerCase();
            const match = keys.find(k => k.toLowerCase() === t);
            if (match && row[match] !== undefined && row[match] !== null && String(row[match]).trim() !== '') return row[match];
          }
          for (const term of searchTerms) {
            const t = term.toLowerCase();
            const match = keys.find(k => k.toLowerCase().includes(t));
            if (match && row[match] !== undefined && row[match] !== null && String(row[match]).trim() !== '') return row[match];
          }
          return null;
        }
      };

      const startIndex = isArrayRows ? 1 : 0;

      // First pass: Identify states, ASINs, Ad Group -> Campaign mapping, and Bidding Adjustments (Placements)
      for (let i = startIndex; i < data.length; i++) {
        const row = data[i];
        if (!row || (isArrayRows && row.length === 0)) continue;

        const entityType = String(getVal(row, ['Entity Type', 'Record Type', 'Entity']) || '').toLowerCase();
        const campaignId = String(getVal(row, ['Campaign ID', 'Campaign Name (Informational only)', 'Campaign Name', 'Campaign']) || '');
        const campaignName = String(getVal(row, ['Campaign Name', 'Campaign Name (Informational only)', 'Campaign', 'Campaign ID']) || '');
        const adGroupId = getVal(row, ['Ad Group ID', 'Ad Group Name', 'Ad Group']);
        const state = String(getVal(row, ['State', 'Status']) || '').toLowerCase();
        const asin = getVal(row, ['Product ID', 'SKU','ASIN (Informational only)', 'ASIN']);
        const placement = getVal(row, ['Placement', 'Placement Type']);
        const percentage = getVal(row, ['Percentage', 'Bid Adjustment', 'Adjustment']);
        const audienceId = getVal(row, ['Audience ID', 'Audience Name', 'Audience']);

        if (entityType === 'campaign' && campaignId) {
          campaignStates[campaignId] = state;
        }
        if (entityType === 'ad group' && adGroupId && campaignId) {
          adGroupStates[adGroupId] = state;
          adGroupToCampaign[adGroupId] = campaignId;
          if (audienceId) adGroupAudiences[adGroupId] = String(audienceId);
        }
        if (entityType === 'bidding adjustment' && campaignId && placement) {
          if (!campaignPlacementProfiles[campaignId]) {
            campaignPlacementProfiles[campaignId] = {};
          }
          // Normalize placement names (e.g., "Placement Top" -> "Top")
          const normPlacement = String(placement).replace(/^Placement\s+/i, '').trim();
          campaignPlacementProfiles[campaignId][normPlacement] = String(percentage || '0');
        }
        if (entityType.includes('product ad') && adGroupId && asin) {
          if (!adGroupAsins[adGroupId]) adGroupAsins[adGroupId] = new Set();
          adGroupAsins[adGroupId].add(String(asin));
        }
      }

      const getPlacementProfileString = (campaignId: string) => {
        const profile = campaignPlacementProfiles[campaignId] || {};
        const sortedPlacements = Object.keys(profile).sort();
        if (sortedPlacements.length === 0) return "Default (0%)";
        return sortedPlacements.map(p => `${p}: ${profile[p]}%`).join(' | ');
      };

      // Second pass: Process keywords
      for (let i = startIndex; i < data.length; i++) {
        const row = data[i];
        if (!row || (isArrayRows && row.length === 0)) continue;

        const entityType = String(getVal(row, ['Entity Type', 'Record Type', 'Entity']) || '').toLowerCase();
        
        if (entityType !== 'keyword') continue;

        const adGroupId = getVal(row, ['Ad Group ID', 'Ad Group Name', 'Ad Group']) || 'Unknown Ad Group';
        const campaignId = String(getVal(row, ['Campaign ID', 'Campaign Name (Informational only)', 'Campaign Name', 'Campaign']) || adGroupToCampaign[adGroupId] || 'Unknown Campaign');
        const campaignName = String(getVal(row, ['Campaign Name', 'Campaign Name (Informational only)', 'Campaign', 'Campaign ID']) || campaignId);
        
        // Step 1: Only enabled campaigns
        const campaignState = campaignStates[campaignId] || 'enabled'; // Default to enabled if not found
        if (campaignState.toLowerCase() !== 'enabled' && campaignState.toLowerCase() !== 'active') continue;

        const keyword = getVal(row, ['Keyword Text', 'Keyword', 'Targeting', 'Product Targeting', 'Targeting Expression', 'Targeting Text']);
        if (!keyword) continue;

        const matchType = getVal(row, ['Match Type', 'Match type']) || 'N/A';
        const rawMatchType = String(matchType).toLowerCase();
        if (rawMatchType.includes('negative')) continue;

        const state = String(getVal(row, ['State', 'Status']) || '').toLowerCase();
        const isKeywordEnabled = (state === 'enabled' || state === 'active' || !state);
        if (!isKeywordEnabled) continue;

        const asinList = adGroupAsins[adGroupId] ? Array.from(adGroupAsins[adGroupId]).sort().join(', ') : 'N/A';
        const audience = adGroupAudiences[adGroupId] || getVal(row, ['Audience ID', 'Audience Name', 'Audience']) || 'N/A';
        const placementProfile = getPlacementProfileString(campaignId);
        
        // Step 4: Placement profile and ASIN are now part of the key
        const key = `${String(keyword).toLowerCase()}|${rawMatchType}|${placementProfile.toLowerCase()}|${String(audience).toLowerCase()}|${asinList.toLowerCase()}`;
        
        if (!groups[key]) {
          groups[key] = {
            keyword: String(keyword),
            matchType: String(matchType),
            asin: asinList,
            placement: placementProfile,
            audience: String(audience),
            instances: []
          };
        }
        
        groups[key].instances.push({
          campaign: campaignName,
          adGroup: String(adGroupId),
          campaignState: campaignState,
          adGroupState: adGroupStates[adGroupId] || 'N/A',
          acos: String(getVal(row, ['ACoS', 'Total ACoS', 'Advertising Cost of Sales']) || '0%'),
          bid: String(getVal(row, ['Bid', 'Max Bid', 'Keyword Bid']) || '0.00'),
          impressions: String(getVal(row, ['Impressions']) || '0'),
          clicks: String(getVal(row, ['Clicks']) || '0'),
          spend: String(getVal(row, ['Spend', 'Advertising Spend']) || '0.00'),
          sales: String(getVal(row, ['Sales', 'Total Sales', '7 Day Total Sales']) || '0.00'),
          orders: String(getVal(row, ['Orders', 'Total Orders', '7 Day Total Orders']) || '0'),
          roas: String(getVal(row, ['ROAS', 'Return on Advertising Spend']) || '0.00'),
          row: i + 1
        });
      }

      const duplicates = Object.values(groups);
      duplicates.sort((a, b) => b.instances.length - a.instances.length);
      setPpcAuditResults(duplicates);
    } catch (err) {
      console.error("PPC Deduplicator failed:", err);
      setPpcError("Failed to process the bulk file. Please ensure it's a valid Amazon bulk file.");
    } finally {
      setIsAuditingPpc(false);
    }
  };

  const auditPpcFile = (file: File) => {
    setIsAuditingPpc(true);
    setPpcError(null);
    setPpcFileName(file.name);

    const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.xlsb');

    if (isExcel) {
      if (file.size > 100 * 1024 * 1024) {
        setPpcError("The Excel file is too large (Max 100MB). Please convert it to CSV or split it into smaller parts for better performance.");
        setIsAuditingPpc(false);
        return;
      }

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          if (!data) throw new Error("Could not read file data");

          // Two-pass read to save memory
          // 1. Get sheet names first
          const tempWorkbook = XLSX.read(data, { type: 'array', bookSheets: true });
          const sheetNames = tempWorkbook.SheetNames;
          
          const targetSheetName = "Sponsored Products Campaigns";
          const foundTargetSheet = sheetNames.find(n => n.toLowerCase() === targetSheetName.toLowerCase());

          // 2. Read only necessary sheets
          const readOptions: any = { 
            type: 'array',
            cellStyles: false,
            cellFormula: false,
            cellHTML: false,
            cellDates: true,
            dense: true
          };

          // Always prioritize the Sponsored Products sheet for both PPC tools
          if (foundTargetSheet) {
            readOptions.sheets = [foundTargetSheet];
          }

          const workbook = XLSX.read(data, readOptions);
          
          let combinedData: any[] = [];
          
          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) return;

            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
            if (Array.isArray(jsonData) && jsonData.length > 0) {
              // If we already have data, skip the header row of subsequent sheets
              const startIdx = combinedData.length > 0 ? 1 : 0;
              for (let i = startIdx; i < jsonData.length; i++) {
                const row = jsonData[i] as any[];
                if (row && row.length > 0) {
                  combinedData.push(row);
                }
              }
            }
          });

          if (combinedData.length === 0) {
            if (foundTargetSheet) {
              setPpcError(`The "${targetSheetName}" tab was found but appears to be empty.`);
            } else {
              setPpcError(`Could not find the "${targetSheetName}" tab in the Excel file. Please ensure you are uploading a standard Amazon Bulk File.`);
            }
            setIsAuditingPpc(false);
            return;
          }
          
          processPpcData(combinedData, true);
        } catch (err: any) {
          console.error("Excel Parsing failed:", err);
          const errorMessage = err?.message || String(err);
          if (errorMessage.includes("allocation") || errorMessage.includes("memory") || errorMessage.includes("size")) {
            setPpcError("Memory limit exceeded. This file is too large for browser processing. Please convert it to CSV or split it into smaller parts.");
          } else {
            setPpcError("Failed to parse the Excel file. Try converting it to CSV.");
          }
          setIsAuditingPpc(false);
        }
      };
      reader.onerror = () => {
        setPpcError("Failed to read the file from disk.");
        setIsAuditingPpc(false);
      };
      reader.readAsArrayBuffer(file);
    } else {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          processPpcData(results.data);
        },
        error: (err) => {
          console.error("CSV Parsing failed:", err);
          setPpcError("Failed to parse the CSV file.");
          setIsAuditingPpc(false);
        }
      });
    }
  };

  return (
    <div className="min-h-screen bg-[#F9FAFB] text-[#111827] font-sans p-4 md:p-8">
      <div className="max-w-5xl mx-auto">
        {/* Header */}
        <header className="mb-8 text-center md:text-left flex flex-col md:flex-row md:items-end justify-between gap-4">
          <div>
            <h1 className="text-3xl font-bold tracking-tight text-gray-900 sm:text-4xl">
              MAG Deduplicator Tool
            </h1>
            <p className="mt-2 text-lg text-gray-600">
              PPC bulk file auditing and deduplication.
            </p>
          </div>
        </header>

        <div className="max-w-5xl mx-auto animate-in fade-in duration-300">
          <div className="space-y-8">
            {/* Upload Section */}
            <div className="bg-white rounded-2xl p-8 shadow-sm border border-gray-100">
              <h2 className="text-xl font-bold text-gray-900 mb-6 flex items-center gap-2">
                <FileSpreadsheet className="w-6 h-6 text-blue-500" />
                Bulk File Upload
              </h2>
              
              <PpcUploadZone 
                fileName={ppcFileName}
                isAuditing={isAuditingPpc}
                onUpload={auditPpcFile}
                onReset={resetPpcAudit}
              />

              <div className="mt-8 p-4 bg-blue-50 rounded-xl border border-blue-100">
                <h3 className="text-sm font-bold text-blue-800 mb-2 flex items-center gap-2">
                  <AlertCircle className="w-4 h-4" />
                  Audit Logic
                </h3>
                <p className="text-xs text-blue-700 leading-relaxed">
                  This tool identifies duplicate targeting by checking <strong>Enabled</strong> keywords in <strong>Enabled</strong> campaigns. It groups keywords by their text, match type, Audience, and <strong>Campaign Placement Profiles</strong> (extracted from Bidding Adjustment rows).
                </p>
              </div>

              {/* Results Section - Now inside the same container below logic */}
              <div className="mt-8 pt-8 border-t border-gray-100">
                <PpcAuditResults 
                  isAuditing={isAuditingPpc}
                  results={ppcAuditResults}
                  error={ppcError}
                  onNewAudit={resetPpcAudit}
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function PpcUploadZone({ fileName, isAuditing, onUpload, onReset }: any) {
  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles[0]) onUpload(acceptedFiles[0]);
  }, [onUpload]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: { 'text/csv': ['.csv'], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
    multiple: false,
    disabled: isAuditing
  } as any);

  return (
    <div 
      {...getRootProps()} 
      className={cn(
        "relative border-2 border-dashed rounded-2xl p-8 transition-all duration-200 cursor-pointer flex flex-col items-center justify-center min-h-[300px]",
        isDragActive ? "border-blue-500 bg-blue-50" : "border-gray-300 bg-white hover:border-gray-400",
        isAuditing && "opacity-50 cursor-not-allowed"
      )}
    >
      <input {...getInputProps()} />
      
      {fileName ? (
        <div className="text-center">
          <div className="mx-auto w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mb-4">
            <FileSpreadsheet className="w-8 h-8 text-blue-500" />
          </div>
          <p className="text-lg font-medium text-gray-900 truncate max-w-[250px]">{fileName}</p>
          <button 
            onClick={(e) => {
              e.stopPropagation();
              onReset();
            }}
            className="mt-4 px-4 py-2 text-sm font-medium text-gray-600 hover:text-gray-900 flex items-center gap-2 mx-auto"
          >
            <RefreshCw className="w-4 h-4" />
            Clear File
          </button>
        </div>
      ) : (
        <div className="text-center">
          <div className="mx-auto w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mb-4">
            <Upload className="w-8 h-8 text-gray-400" />
          </div>
          <p className="text-lg font-medium text-gray-900">Upload Amazon Bulk File</p>
          <p className="text-sm text-gray-500 mt-1">Drag & drop CSV or click to browse</p>
        </div>
      )}
      
      {isAuditing && (
        <div className="absolute inset-0 bg-white/80 backdrop-blur-sm rounded-2xl flex flex-col items-center justify-center">
          <Loader2 className="w-10 h-10 text-blue-500 animate-spin mb-2" />
          <p className="text-sm font-bold text-gray-600">Analyzing Bulk Data...</p>
        </div>
      )}
    </div>
  );
}

function PpcAuditResults({ isAuditing, results, error, onNewAudit }: any) {
  const [searchTerm, setSearchTerm] = useState('');
  const [filter, setFilter] = useState<'all' | 'duplicates'>('all');

  if (isAuditing) return null;

  if (error) {
    return (
      <div className="bg-red-50 rounded-2xl p-8 border border-red-100 text-center">
        <AlertCircle className="w-12 h-12 text-red-500 mx-auto mb-4" />
        <h3 className="text-lg font-semibold text-red-900">Audit Error</h3>
        <p className="text-red-700 mt-2 mb-6">{error}</p>
        <button 
          onClick={onNewAudit}
          className="px-6 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors font-medium"
        >
          Try Again
        </button>
      </div>
    );
  }

  if (!results) {
    return (
      <div className="bg-white rounded-2xl p-12 shadow-sm border border-gray-100 flex flex-col items-center justify-center text-center min-h-[400px]">
        <Search className="w-16 h-16 text-gray-200 mb-4" />
        <h3 className="text-xl font-semibold text-gray-400">No PPC Deduplicator Results</h3>
        <p className="text-gray-400 mt-2 max-w-xs">
          Upload a bulk file to identify duplicate keyword targeting.
        </p>
      </div>
    );
  }

  const duplicateGroups = results.filter((g: any) => {
    const adGroupIds = new Set(g.instances.map((i: any) => i.adGroup.toLowerCase()));
    return adGroupIds.size > 1;
  });
  const duplicateGroupsCount = duplicateGroups.length;
  
  const filteredResults = results.filter((g: any) => {
    const matchesSearch = g.keyword.toLowerCase().includes(searchTerm.toLowerCase());
    const adGroupIds = new Set(g.instances.map((i: any) => i.adGroup.toLowerCase()));
    const matchesFilter = filter === 'all' || adGroupIds.size > 1;
    return matchesSearch && matchesFilter;
  });

  const getCampaignCount = (group: PpcDuplicateGroup) => {
    return new Set(group.instances.map(i => i.campaign.toLowerCase())).size;
  };

  const getAdGroupCount = (group: PpcDuplicateGroup) => {
    return new Set(group.instances.map(i => i.adGroup.toLowerCase())).size;
  };

  return (
    <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <div className="flex items-center justify-between">
        <h3 className="text-sm font-semibold text-gray-500 uppercase tracking-wider">PPC Deduplicator Report</h3>
        <button 
          onClick={onNewAudit}
          className="text-sm font-medium text-blue-600 hover:text-blue-700 flex items-center gap-1"
        >
          <RefreshCw className="w-3 h-3" />
          New Audit
        </button>
      </div>

      <div className={cn(
        "rounded-2xl p-6 border flex items-center gap-4",
        duplicateGroupsCount === 0 
          ? "bg-green-50 border-green-100 text-green-800" 
          : "bg-red-50 border-red-100 text-red-800"
      )}>
        {duplicateGroupsCount === 0 ? (
          <CheckCircle2 className="w-10 h-10 text-green-500 shrink-0" />
        ) : (
          <XCircle className="w-10 h-10 text-red-500 shrink-0" />
        )}
        <div>
          <h3 className="text-xl font-bold">
            {duplicateGroupsCount === 0 ? "No Duplicates Found" : `${duplicateGroupsCount} Duplicate Keyword Groups Found`}
          </h3>
          <p className="text-sm opacity-90">
            {duplicateGroupsCount === 0 
              ? "Your campaign structure is clean and efficient." 
              : `We found ${duplicateGroupsCount} groups of keywords targeted across different Ad Groups or Campaigns with the same Match Type, Audience, and Bidding Adjustment (Placement) settings.`}
          </p>
        </div>
      </div>

      {/* Search and Filter */}
      <div className="flex flex-col sm:flex-row gap-4">
        <div className="relative flex-1">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
          <input 
            type="text" 
            placeholder="Search keywords..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all text-sm"
          />
        </div>
        <div className="flex bg-gray-100 p-1 rounded-xl border border-gray-200">
          <button 
            onClick={() => setFilter('all')}
            className={cn(
              "px-4 py-1.5 text-xs font-bold rounded-lg transition-all",
              filter === 'all' ? "bg-white text-blue-600 shadow-sm" : "text-gray-500 hover:text-gray-700"
            )}
          >
            All ({results.length})
          </button>
          <button 
            onClick={() => setFilter('duplicates')}
            className={cn(
              "px-4 py-1.5 text-xs font-bold rounded-lg transition-all",
              filter === 'duplicates' ? "bg-white text-red-600 shadow-sm" : "text-gray-500 hover:text-gray-700"
            )}
          >
            Duplicates ({duplicateGroupsCount})
          </button>
        </div>
      </div>

      {filteredResults.length > 0 ? (
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50 border-b border-gray-100">
                  <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">
                    {filter === 'duplicates' ? 'Keyword & Campaign' : 'Keyword'}
                  </th>
                  <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">
                    {filter === 'duplicates' ? 'Ad Group ID' : 'Ad Group'}
                  </th>
                  {filter === 'duplicates' && (
                    <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Advertised ASIN</th>
                  )}
                  <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Match Type</th>
                  <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Placement</th>
                  <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Audience</th>
                  {filter === 'duplicates' && (
                    <>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Bid</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Impr.</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Clicks</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Spend</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Sales</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">Orders</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">ROAS</th>
                      <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400">ACoS</th>
                    </>
                  )}
                  <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-gray-400 text-center">
                    {filter === 'duplicates' ? 'Row' : 'Instances'}
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                  {filteredResults.flatMap((group: PpcDuplicateGroup, groupIdx: number) => {
                    const adGroupCount = getAdGroupCount(group);
                    const instancesToShow = filter === 'duplicates' ? group.instances : [group.instances[0]];
                    
                    return instancesToShow.map((instance, instanceIdx) => (
                      <tr key={`${groupIdx}-${instanceIdx}`} className="hover:bg-gray-50/50 transition-colors">
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-3">
                            <div className={cn(
                              "p-1.5 rounded-lg shrink-0",
                              adGroupCount > 1 ? "bg-red-100 text-red-600" : "bg-blue-100 text-blue-600"
                            )}>
                              <Target className="w-3.5 h-3.5" />
                            </div>
                            <div className="flex flex-col">
                              <span className="font-bold text-gray-900 text-sm">{group.keyword}</span>
                              {filter === 'duplicates' ? (
                                <span className="text-[10px] text-gray-400 font-medium truncate max-w-[150px]">
                                  {instance.campaign}
                                </span>
                              ) : (
                                adGroupCount > 1 && (
                                  <span className="text-[9px] text-red-500 font-bold uppercase tracking-tight">
                                    Targeted in {getCampaignCount(group)} Campaigns
                                  </span>
                                )
                              )}
                            </div>
                          </div>
                        </td>
                      <td className="px-6 py-4">
                        <div className="flex flex-col">
                          <span className="text-gray-900 text-xs font-medium">
                            {filter === 'duplicates' ? instance.adGroup : (adGroupCount > 1 ? 'Multiple Ad Groups' : instance.adGroup)}
                          </span>
                          {(filter === 'duplicates' || adGroupCount === 1) && (
                            <span className={cn(
                              "text-[9px] font-bold uppercase tracking-wider",
                              instance.adGroupState === 'enabled' ? "text-green-500" : "text-gray-400"
                            )}>
                              {instance.adGroupState}
                            </span>
                          )}
                        </div>
                      </td>
                      {filter === 'duplicates' && (
                        <td className="px-6 py-4">
                          <span className="text-gray-600 text-xs font-mono">{group.asin}</span>
                        </td>
                      )}
                      <td className="px-6 py-4">
                        <span className="px-2 py-0.5 bg-gray-100 text-gray-600 rounded text-[10px] font-bold uppercase">{group.matchType}</span>
                      </td>
                      <td className="px-6 py-4">
                        <span className="text-gray-600 text-xs">{group.placement}</span>
                      </td>
                      <td className="px-6 py-4">
                        <span className="text-gray-600 text-xs">{group.audience}</span>
                      </td>
                      {filter === 'duplicates' && (
                        <>
                          <td className="px-6 py-4">
                            <span className="text-gray-900 text-xs font-medium">{instance.bid}</span>
                          </td>
                          <td className="px-6 py-4">
                            <span className="text-gray-600 text-xs">{instance.impressions}</span>
                          </td>
                          <td className="px-6 py-4">
                            <span className="text-gray-600 text-xs">{instance.clicks}</span>
                          </td>
                          <td className="px-6 py-4">
                            <span className="text-gray-900 text-xs font-medium">${instance.spend}</span>
                          </td>
                          <td className="px-6 py-4">
                            <span className="text-gray-900 text-xs font-medium">${instance.sales}</span>
                          </td>
                          <td className="px-6 py-4">
                            <span className="text-gray-600 text-xs">{instance.orders}</span>
                          </td>
                          <td className="px-6 py-4">
                            <span className="text-gray-900 text-xs font-medium">{instance.roas}</span>
                          </td>
                          <td className="px-6 py-4">
                            <span className="text-gray-900 text-xs font-medium">{instance.acos}</span>
                          </td>
                        </>
                      )}
                      <td className="px-6 py-4 text-center">
                        {filter === 'duplicates' ? (
                          <span className="text-gray-400 text-[10px] font-mono">#{instance.row}</span>
                        ) : (
                          <div className="flex flex-col items-center gap-1">
                            <div className={cn(
                              "inline-flex items-center justify-center min-w-[24px] h-6 px-2 rounded-full text-xs font-bold",
                              adGroupCount > 1 ? "bg-red-100 text-red-700" : "bg-blue-100 text-blue-700"
                            )}>
                              {adGroupCount}
                            </div>
                            {adGroupCount > 1 && (
                              <span className="text-[9px] font-bold text-red-500 uppercase whitespace-nowrap">
                                {adGroupCount - 1} Duplicates
                              </span>
                            )}
                          </div>
                        )}
                      </td>
                    </tr>
                  ));
                })}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div className="bg-white rounded-2xl p-12 shadow-sm border border-gray-100 text-center">
          <Search className="w-12 h-12 text-gray-200 mx-auto mb-4" />
          <p className="text-gray-500">No keywords match your search or filter.</p>
        </div>
      )}
    </div>
  );
}


