import { useState, useEffect, useRef, useCallback } from "react";

const WIN1252_MAP = {0x80:0x20AC,0x82:0x201A,0x83:0x0192,0x84:0x201E,0x85:0x2026,0x86:0x2020,0x87:0x2021,0x88:0x02C6,0x89:0x2030,0x8A:0x0160,0x8B:0x2039,0x8C:0x0152,0x8E:0x017D,0x91:0x2018,0x92:0x2019,0x93:0x201C,0x94:0x201D,0x95:0x2022,0x96:0x2013,0x97:0x2014,0x98:0x02DC,0x99:0x2122,0x9A:0x0161,0x9B:0x203A,0x9C:0x0153,0x9E:0x017E,0x9F:0x0178};

function detectAndDecode(buffer) {
  const b = new Uint8Array(buffer);
  let enc = "utf-8", off = 0;
  if (b[0]===0xEF&&b[1]===0xBB&&b[2]===0xBF) { enc="utf-8"; off=3; }
  else if (b[0]===0xFF&&b[1]===0xFE) { enc="utf-16le"; off=2; }
  else if (b[0]===0xFE&&b[1]===0xFF) { enc="utf-16be"; off=2; }
  else {
    const head = new TextDecoder("ascii",{fatal:false}).decode(b.slice(0,2048));
    const m = head.match(/charset\s*=\s*["']?\s*([\w-]+)/i);
    if (m) {
      const cs = m[1].toLowerCase();
      if (cs.includes("1252")||cs.includes("windows")) enc = "windows-1252";
      else if (cs.includes("8859")) enc = "iso-8859-1";
    } else {
      let w=false, inv=false;
      for (let i=0;i<Math.min(b.length,8192);i++) {
        if (b[i]>=0x80&&b[i]<=0x9F) w=true;
        if (b[i]>=0xC0&&b[i]<=0xDF&&(i+1>=b.length||(b[i+1]&0xC0)!==0x80)) inv=true;
      }
      if (w) enc="windows-1252"; else if (inv) enc="iso-8859-1";
    }
  }
  if (enc==="windows-1252") {
    let r=""; const bytes=new Uint8Array(buffer,off);
    for (let i=0;i<bytes.length;i++) { const c=bytes[i]; r+=c<0x80?String.fromCharCode(c):String.fromCharCode(WIN1252_MAP[c]||c); }
    return { text: r, encoding: enc };
  }
  try { return { text: new TextDecoder(enc).decode(new Uint8Array(buffer,off)), encoding: enc }; }
  catch { return { text: new TextDecoder("utf-8",{fatal:false}).decode(b), encoding: "utf-8 (fallback)" }; }
}

const VARS = [
  { token:"{{DisplayName}}", label:"Display Name", sample:"Max Mustermann" },
  { token:"{{GivenName}}", label:"First Name", sample:"Max" },
  { token:"{{Surname}}", label:"Last Name", sample:"Mustermann" },
  { token:"{{JobTitle}}", label:"Job Title", sample:"Senior Consultant" },
  { token:"{{Department}}", label:"Department", sample:"IT Infrastructure" },
  { token:"{{Mail}}", label:"Email", sample:"m.mustermann@datagroup.de" },
  { token:"{{Phone}}", label:"Phone", sample:"+49 711 123456-0" },
  { token:"{{Mobile}}", label:"Mobile", sample:"+49 170 1234567" },
  { token:"{{Company}}", label:"Company", sample:"DATAGROUP SE" },
  { token:"{{Office}}", label:"Office", sample:"Stuttgart HQ" },
  { token:"{{Street}}", label:"Street", sample:"Wilhelm-Schickard-Str. 7" },
  { token:"{{City}}", label:"City", sample:"Stuttgart" },
  { token:"{{PostalCode}}", label:"Postal Code", sample:"70563" },
  { token:"{{Country}}", label:"Country", sample:"Germany" },
  { token:"{{State}}", label:"State", sample:"Baden-Württemberg" },
  { token:"{{ManagerName}}", label:"Manager", sample:"Erika Beispiel" },
  { token:"{{ExtAttr1}}", label:"ExtAttr 1", sample:"" },
];

const DEFAULT_HTML = `<table cellpadding="0" cellspacing="0" border="0" style="font-family:'Segoe UI',Calibri,Arial,Helvetica,sans-serif;font-size:10pt;color:#333333;max-width:600px;">
  <tr>
    <td style="padding-right:15px;border-right:3px solid #E30613;vertical-align:top;">
      <img src="https://via.placeholder.com/120x45/E30613/FFFFFF?text=LOGO" alt="Logo" width="120" height="45" style="display:block;border:0;" />
    </td>
    <td style="padding-left:15px;vertical-align:top;">
      <table cellpadding="0" cellspacing="0" border="0">
        <tr><td style="font-size:13pt;font-weight:600;color:#1a1a1a;padding-bottom:2px;">{{DisplayName}}</td></tr>
        <tr><td style="font-size:9pt;color:#E30613;text-transform:uppercase;letter-spacing:0.5px;padding-bottom:8px;">{{JobTitle}} · {{Department}}</td></tr>
        <tr><td style="font-size:9pt;color:#555555;line-height:1.7;">
          <span style="color:#888;">T</span>&nbsp;{{Phone}}<br/>
          <span style="color:#888;">M</span>&nbsp;{{Mobile}}<br/>
          <span style="color:#888;">E</span>&nbsp;<a href="mailto:{{Mail}}" style="color:#0078D4;text-decoration:none;">{{Mail}}</a>
        </td></tr>
        <tr><td style="padding-top:8px;font-size:8pt;color:#999999;border-top:1px solid #e0e0e0;">
          {{Company}} · {{Street}} · {{PostalCode}} {{City}}
        </td></tr>
      </table>
    </td>
  </tr>
</table>`;

function expandVars(html) {
  let r = html;
  VARS.forEach(v => { r = r.replaceAll(v.token, v.sample || ""); });
  r = r.replace(/\{\{\w+\}\}/g, "");
  return r;
}

function VarPill({ v, onClick }) {
  return (
    <button onClick={() => onClick(v.token)}
      style={{ display:"inline-flex", alignItems:"center", gap:4, padding:"3px 10px", fontSize:12, borderRadius:4,
        background:"var(--color-background-info,#E6F1FB)", color:"var(--color-text-info,#185FA5)",
        border:"0.5px solid var(--color-border-info,#85B7EB)", cursor:"pointer", whiteSpace:"nowrap" }}
      title={`Inserts ${v.token} → "${v.sample}"`}>
      <span style={{fontFamily:"var(--font-mono,monospace)",fontSize:11}}>{v.token}</span>
    </button>
  );
}

function TabBtn({ active, children, onClick }) {
  return (
    <button onClick={onClick}
      style={{ padding:"8px 18px", fontSize:13, fontWeight:active?500:400, cursor:"pointer",
        borderBottom:active?"2px solid var(--color-text-info,#185FA5)":"2px solid transparent",
        color:active?"var(--color-text-primary)":"var(--color-text-secondary)",
        background:"none", border:"none", borderBottomWidth:2, borderBottomStyle:"solid",
        borderBottomColor:active?"var(--color-text-info,#185FA5)":"transparent" }}>
      {children}
    </button>
  );
}

export default function SignatureManager() {
  const [html, setHtml] = useState(DEFAULT_HTML);
  const [tab, setTab] = useState("editor");
  const [fileInfo, setFileInfo] = useState(null);
  const [copied, setCopied] = useState(false);
  const [psTab, setPsTab] = useState("roaming");
  const editorRef = useRef(null);
  const textareaRef = useRef(null);
  const fileInputRef = useRef(null);

  const insertVar = useCallback((token) => {
    if (tab === "source") {
      const ta = textareaRef.current;
      if (ta) {
        const start = ta.selectionStart, end = ta.selectionEnd;
        const newHtml = html.slice(0, start) + token + html.slice(end);
        setHtml(newHtml);
        setTimeout(() => { ta.selectionStart = ta.selectionEnd = start + token.length; ta.focus(); }, 0);
      }
    } else {
      setHtml(prev => prev + token);
    }
  }, [tab, html]);

  const handleFile = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const { text, encoding } = detectAndDecode(buffer);
    setHtml(text);
    setFileInfo({ name: file.name, size: file.size, encoding });
    setTab("source");
  }, []);

  const exportHtml = useCallback(() => {
    const fullHtml = `<!DOCTYPE html>\n<html lang="de">\n<head>\n<meta charset="utf-8">\n<title>Email Signature</title>\n</head>\n<body style="margin:0;padding:0;">\n${html}\n</body>\n</html>`;
    const blob = new Blob([new Uint8Array([0xEF,0xBB,0xBF]), fullHtml], { type:"text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = "signature.htm"; a.click();
    URL.revokeObjectURL(url);
  }, [html]);

  const copyHtml = useCallback(() => {
    navigator.clipboard.writeText(html).then(() => { setCopied(true); setTimeout(() => setCopied(false), 2000); });
  }, [html]);

  const psCommands = {
    roaming: `# Strategy: Roaming (Graph Beta UserConfiguration API)
# Requires: MailboxConfigItem.ReadWrite, Mail.ReadWrite, User.Read.All
.\\Manage-OutlookSignatures.ps1 \`
    -TenantId "YOUR-TENANT-ID" \`
    -ClientId "YOUR-CLIENT-ID" \`
    -TemplatePath ".\\templates\\corporate.htm" \`
    -UserUPN "user@contoso.com" \`
    -Strategy Roaming \`
    -SetAsDefault -SetForReply`,
    transport: `# Strategy: Transport Rule (server-side, ALL clients)
# Requires: Exchange.ManageAsApp, ClientSecret
# Note: {{Variables}} auto-convert to %%exchangeVars%%
.\\Manage-OutlookSignatures.ps1 \`
    -TenantId "YOUR-TENANT-ID" \`
    -ClientId "YOUR-CLIENT-ID" \`
    -ClientSecret "YOUR-SECRET" \`
    -TemplatePath ".\\templates\\corporate.htm" \`
    -Strategy TransportRule`,
    bulk: `# Bulk deploy to all licensed users (app-only)
# Requires: ClientSecret + all permissions
.\\Manage-OutlookSignatures.ps1 \`
    -TenantId "YOUR-TENANT-ID" \`
    -ClientId "YOUR-CLIENT-ID" \`
    -ClientSecret "YOUR-SECRET" \`
    -TemplatePath ".\\templates\\corporate.htm" \`
    -UserUPN "*" \`
    -Strategy Roaming \`
    -SetAsDefault -SetForReply`,
    appReg: `# Create Entra ID App Registration
# Run once, requires Global Admin or App Admin
.\\Register-SignatureManagerApp.ps1 \`
    -AppName "Signature Manager" \`
    -CreateClientSecret

# Required Graph permissions:
# - MailboxConfigItem.ReadWrite  (roaming sigs)
# - Mail.ReadWrite               (mailbox access)
# - User.Read.All                (user properties)
# - MailboxSettings.ReadWrite    (OOF replies)
# - Exchange.ManageAsApp         (transport rules)`
  };

  const previewHtml = expandVars(html);
  const varCount = (html.match(/\{\{\w+\}\}/g) || []).length;
  const sizeKb = new Blob([html]).size / 1024;

  return (
    <div style={{ fontFamily:"var(--font-sans, 'Segoe UI', system-ui, sans-serif)", color:"var(--color-text-primary, #1a1a1a)" }}>
      {/* Header bar */}
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", gap:12, marginBottom:16, flexWrap:"wrap" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:32, height:32, borderRadius:"var(--border-radius-md,8px)", background:"#E30613", display:"flex", alignItems:"center", justifyContent:"center" }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="M22 7L12 13L2 7"/></svg>
          </div>
          <span style={{ fontSize:16, fontWeight:500 }}>Signature designer</span>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:12, color:"var(--color-text-secondary)", padding:"4px 10px", background:"var(--color-background-secondary)", borderRadius:4 }}>
            {varCount} variables · {sizeKb.toFixed(1)} KB
          </span>
          <input ref={fileInputRef} type="file" accept=".htm,.html,.txt" onChange={handleFile} style={{display:"none"}} />
          <button onClick={() => fileInputRef.current?.click()} style={{ padding:"6px 14px", fontSize:12, cursor:"pointer", borderRadius:4, border:"0.5px solid var(--color-border-secondary)", background:"var(--color-background-primary)", color:"var(--color-text-primary)" }}>
            Open file
          </button>
          <button onClick={exportHtml} style={{ padding:"6px 14px", fontSize:12, cursor:"pointer", borderRadius:4, border:"0.5px solid var(--color-border-secondary)", background:"var(--color-background-primary)", color:"var(--color-text-primary)" }}>
            Export .htm
          </button>
          <button onClick={copyHtml} style={{ padding:"6px 14px", fontSize:12, cursor:"pointer", borderRadius:4, border:"0.5px solid var(--color-border-info)", background:"var(--color-background-info)", color:"var(--color-text-info)" }}>
            {copied ? "Copied!" : "Copy HTML"}
          </button>
        </div>
      </div>

      {/* File info banner */}
      {fileInfo && (
        <div style={{ fontSize:12, color:"var(--color-text-secondary)", padding:"6px 12px", marginBottom:12,
          background:"var(--color-background-secondary)", borderRadius:4, display:"flex", gap:16, alignItems:"center" }}>
          <span>Loaded: <strong>{fileInfo.name}</strong></span>
          <span>{(fileInfo.size/1024).toFixed(1)} KB</span>
          <span style={{ padding:"2px 8px", borderRadius:3, background:"var(--color-background-info)", color:"var(--color-text-info)" }}>
            {fileInfo.encoding}
          </span>
          <button onClick={() => setFileInfo(null)} style={{ marginLeft:"auto", cursor:"pointer", background:"none", border:"none", color:"var(--color-text-secondary)", fontSize:12 }}>dismiss</button>
        </div>
      )}

      {/* Variable insertion toolbar */}
      <div style={{ marginBottom:12 }}>
        <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:6 }}>Insert variable (click to add at cursor in source view):</div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
          {VARS.map(v => <VarPill key={v.token} v={v} onClick={insertVar} />)}
        </div>
      </div>

      {/* Tab bar */}
      <div style={{ display:"flex", borderBottom:"0.5px solid var(--color-border-tertiary)", marginBottom:0 }}>
        <TabBtn active={tab==="editor"} onClick={() => setTab("editor")}>Visual preview</TabBtn>
        <TabBtn active={tab==="source"} onClick={() => setTab("source")}>HTML source</TabBtn>
        <TabBtn active={tab==="preview"} onClick={() => setTab("preview")}>Live preview (with data)</TabBtn>
        <TabBtn active={tab==="deploy"} onClick={() => setTab("deploy")}>Deploy</TabBtn>
      </div>

      {/* Tab content */}
      <div style={{ border:"0.5px solid var(--color-border-tertiary)", borderTop:"none", borderRadius:"0 0 var(--border-radius-md,8px) var(--border-radius-md,8px)", minHeight:340 }}>

        {tab === "editor" && (
          <div style={{ padding:20 }}>
            <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:12 }}>
              Raw template with variable tokens visible. Edit in the HTML source tab for full control.
            </div>
            <div style={{ background:"white", border:"0.5px solid var(--color-border-tertiary)", borderRadius:4, padding:16, overflow:"auto" }}
              dangerouslySetInnerHTML={{ __html: html }} />
          </div>
        )}

        {tab === "source" && (
          <div style={{ padding:20 }}>
            <textarea ref={textareaRef} value={html} onChange={e => setHtml(e.target.value)}
              spellCheck={false}
              style={{ width:"100%", minHeight:300, fontFamily:"var(--font-mono, 'Cascadia Code', 'Fira Code', monospace)",
                fontSize:12, lineHeight:1.6, padding:12, border:"0.5px solid var(--color-border-tertiary)", borderRadius:4,
                background:"var(--color-background-secondary)", color:"var(--color-text-primary)", resize:"vertical",
                tabSize:2 }} />
          </div>
        )}

        {tab === "preview" && (
          <div style={{ padding:20 }}>
            <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:12 }}>
              Variables replaced with sample data. This is how the signature will look for a user.
            </div>
            <div style={{ background:"white", border:"0.5px solid var(--color-border-tertiary)", borderRadius:4, padding:20, overflow:"auto" }}>
              <div style={{ borderBottom:"0.5px solid #e0e0e0", paddingBottom:12, marginBottom:12, fontSize:12, color:"#888" }}>
                <strong style={{ color:"#333" }}>From:</strong> Max Mustermann &lt;m.mustermann@datagroup.de&gt;<br/>
                <strong style={{ color:"#333" }}>To:</strong> kunde@example.com<br/>
                <strong style={{ color:"#333" }}>Subject:</strong> Angebot — Managed Services
              </div>
              <div style={{ fontSize:"10pt", color:"#333", fontFamily:"'Segoe UI',Calibri,Arial,sans-serif", lineHeight:1.5, marginBottom:16 }}>
                <p>Sehr geehrte Damen und Herren,</p>
                <p>anbei erhalten Sie das gewünschte Angebot...</p>
                <p>Mit freundlichen Grüßen</p>
              </div>
              <div dangerouslySetInnerHTML={{ __html: previewHtml }} />
            </div>
          </div>
        )}

        {tab === "deploy" && (
          <div style={{ padding:20 }}>
            <div style={{ fontSize:13, color:"var(--color-text-secondary)", marginBottom:16, lineHeight:1.6 }}>
              Generate the PowerShell command for your deployment scenario. The scripts use the
              Graph Beta UserConfiguration API (<span style={{fontFamily:"var(--font-mono)",fontSize:12}}>MailboxConfigItem.ReadWrite</span>) for
              roaming signatures — no dependency on the deprecated <span style={{fontFamily:"var(--font-mono)",fontSize:12}}>Set-MailboxMessageConfiguration</span>.
            </div>

            <div style={{ display:"flex", gap:6, marginBottom:12, flexWrap:"wrap" }}>
              {[
                { key:"appReg", label:"1. App registration" },
                { key:"roaming", label:"2. Single user" },
                { key:"bulk", label:"3. All users" },
                { key:"transport", label:"4. Transport rule" },
              ].map(t => (
                <button key={t.key} onClick={() => setPsTab(t.key)}
                  style={{ padding:"6px 14px", fontSize:12, cursor:"pointer", borderRadius:4,
                    border: psTab===t.key ? "0.5px solid var(--color-border-info)" : "0.5px solid var(--color-border-tertiary)",
                    background: psTab===t.key ? "var(--color-background-info)" : "var(--color-background-primary)",
                    color: psTab===t.key ? "var(--color-text-info)" : "var(--color-text-primary)" }}>
                  {t.label}
                </button>
              ))}
            </div>

            <pre style={{ fontFamily:"var(--font-mono, monospace)", fontSize:12, lineHeight:1.7, padding:16,
              background:"var(--color-background-secondary)", borderRadius:4, border:"0.5px solid var(--color-border-tertiary)",
              overflow:"auto", whiteSpace:"pre-wrap", wordBreak:"break-word", color:"var(--color-text-primary)" }}>
              {psCommands[psTab]}
            </pre>

            <button onClick={() => navigator.clipboard.writeText(psCommands[psTab])}
              style={{ marginTop:8, padding:"6px 14px", fontSize:12, cursor:"pointer", borderRadius:4,
                border:"0.5px solid var(--color-border-secondary)", background:"var(--color-background-primary)",
                color:"var(--color-text-primary)" }}>
              Copy command
            </button>

            {psTab === "appReg" && (
              <div style={{ marginTop:16, padding:14, borderRadius:4, background:"var(--color-background-warning)", border:"0.5px solid var(--color-border-warning)" }}>
                <div style={{ fontSize:13, fontWeight:500, color:"var(--color-text-warning)", marginBottom:6 }}>Permissions overview</div>
                <table style={{ width:"100%", fontSize:12, borderCollapse:"collapse" }}>
                  <thead>
                    <tr style={{ textAlign:"left", color:"var(--color-text-secondary)" }}>
                      <th style={{ padding:"4px 8px", borderBottom:"0.5px solid var(--color-border-tertiary)" }}>Permission</th>
                      <th style={{ padding:"4px 8px", borderBottom:"0.5px solid var(--color-border-tertiary)" }}>Type</th>
                      <th style={{ padding:"4px 8px", borderBottom:"0.5px solid var(--color-border-tertiary)" }}>Purpose</th>
                    </tr>
                  </thead>
                  <tbody style={{ color:"var(--color-text-primary)" }}>
                    {[
                      ["MailboxConfigItem.ReadWrite", "Delegated+App", "Roaming signatures (UserConfiguration API)"],
                      ["Mail.ReadWrite", "Delegated+App", "Mailbox access"],
                      ["User.Read.All", "Delegated+App", "User properties for template variables"],
                      ["MailboxSettings.ReadWrite", "Delegated+App", "OOF replies and mailbox settings"],
                      ["Exchange.ManageAsApp", "Application", "Transport rules, EXO InvokeCommand"],
                    ].map(([p,t,d],i) => (
                      <tr key={i}>
                        <td style={{ padding:"4px 8px", fontFamily:"var(--font-mono)", fontSize:11, borderBottom:"0.5px solid var(--color-border-tertiary)" }}>{p}</td>
                        <td style={{ padding:"4px 8px", borderBottom:"0.5px solid var(--color-border-tertiary)" }}>{t}</td>
                        <td style={{ padding:"4px 8px", borderBottom:"0.5px solid var(--color-border-tertiary)" }}>{d}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
