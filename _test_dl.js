// Test script: OneDrive download strategies
async function main() {
  const shareUrl = 'https://1drv.ms/x/c/a5c8ae2f1213eda1/IQBsx93bcfE1TK8-08MrXkCZAbfvyfIgOBXqW-zf1oNGlTw?e=FA3lDr';
  
  let cookies = {};
  let viewerUrl = '';
  let url = shareUrl;
  
  // Phase 1: Follow redirects to get cookies + viewer URL
  console.log('=== Phase 1: Follow sharing link ===');
  for (let i = 0; i < 10; i++) {
    const ck = Object.entries(cookies).map(([k,v]) => k + '=' + v).join('; ');
    const r = await fetch(url, {
      redirect: 'manual',
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)', 'Cookie': ck }
    });
    const sc = r.headers.getSetCookie ? r.headers.getSetCookie() : [];
    sc.forEach(c => { const m = c.match(/^([^=]+)=([^;]*)/); if (m) cookies[m[1]] = m[2]; });
    const loc = r.headers.get('location');
    
    if (loc && r.status >= 300 && r.status < 400) {
      url = new URL(loc, url).href;
    } else {
      viewerUrl = url;
      const html = await r.text();
      console.log(`Viewer: ${r.status} (${html.length} bytes) cookies: ${Object.keys(cookies).join(',')}`);
      
      // Look for all potential download patterns
      const patterns = {
        'downloadUrl JSON': /"downloadUrl":"([^"]*)"/,
        '@content.downloadUrl': /"@content\.downloadUrl":"([^"]*)"/,
        'FileSaveUrl': /"FileSaveUrl":"([^"]*)"/,
        'FileGetUrl': /"FileGetUrl":"([^"]*)"/,
        'wopiSrc': /"wopiSrc":"([^"]*)"/,
        'BaseFileName': /"BaseFileName":"([^"]*)"/,
      };
      for (const [name, pat] of Object.entries(patterns)) {
        const m = html.match(pat);
        if (m) console.log(`  ${name}: ${m[1].slice(0,150)}`);
      }
      break;
    }
  }
  
  const cookieStr = Object.entries(cookies).map(([k,v]) => k + '=' + v).join('; ');
  
  // Phase 2: Try various download approaches
  const baseMatch = viewerUrl.match(/(https:\/\/[^/]+\/personal\/[^/]+)/);
  const base = baseMatch ? baseMatch[1] : '';
  
  const strategies = [
    // Strategy A: Viewer URL with action=download
    viewerUrl.replace('action=default', 'action=download').replace('Doc.aspx', 'download.aspx'),
    // Strategy B: _api endpoint
    base + '/_layouts/download.aspx?SourceUrl=' + encodeURIComponent('/personal/a5c8ae2f1213eda1/Documents/SCS_Data.xlsx'),
    // Strategy C: Direct with UniqueId
    base + '/_layouts/15/download.aspx?UniqueId=dbddc76c-f171-4c35-af3e-d3c32b5e4099',
  ];
  
  for (let s = 0; s < strategies.length; s++) {
    const dlUrl = strategies[s];
    console.log(`\n=== Strategy ${s}: ${dlUrl.slice(0,120)}... ===`);
    try {
      const r = await fetch(dlUrl, {
        redirect: 'follow',
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)', 'Cookie': cookieStr }
      });
      const buf = Buffer.from(await r.arrayBuffer());
      const isXlsx = buf[0] === 0x50 && buf[1] === 0x4B;
      console.log(`  Status:${r.status} CT:${r.headers.get('content-type')} Size:${buf.byteLength} XLSX:${isXlsx}`);
      if (isXlsx) {
        console.log('  >>> XLSX FOUND! <<<');
        const XLSX = require('xlsx');
        const wb = XLSX.read(buf, {type:'buffer'});
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:''});
        console.log(`  Sheets:${wb.SheetNames} Rows:${data.length}`);
        if (data[0]) console.log('  Cols:', Object.keys(data[0]).join(', '));
        break;
      }
    } catch(e) { console.log(`  Error: ${e.message}`); }
  }
}
main().catch(e => console.error(e));
