import os
import time
from typing import Dict


def _ensure_index(report_dir: str) -> str:
        os.makedirs(report_dir, exist_ok=True)
        index_path = os.path.join(report_dir, "index.html")
        if not os.path.exists(index_path):
            with open(index_path, "w", encoding="utf-8") as f:
                f.write("""
<!doctype html>
<html>
<head>
    <meta charset='utf-8'>
    <title>DOC Compare Reports</title>
    <style>
        body{font-family:Segoe UI,Arial,sans-serif;margin:20px}
        table{border-collapse:collapse;width:100%}
        th,td{border:1px solid #ccc;padding:6px;text-align:left}
        #controls{margin-bottom:10px}
    </style>
</head>
<body>
    <h1>DOC Compare Reports</h1>
    <div id="controls">
        <input id="searchBox" placeholder="Search JID, AID, summary..." style="width:40%" />
        <button id="prevBtn">Prev</button>
        <span id="pageInfo"></span>
        <button id="nextBtn">Next</button>
        <select id="pageSize">
            <option>10</option>
            <option selected>20</option>
            <option>50</option>
        </select>
    </div>
    <table id="reportTable">
        <thead><tr><th>Timestamp</th><th>JID</th><th>AID</th><th>Summary</th><th>Links</th></tr></thead>
        <tbody id="reportBody">
        </tbody>
    </table>

    <script>
        (function(){
            let allRows = [];
            let currentPage = 1;
            function rebuild(){
                const body = document.getElementById('reportBody');
                // capture any existing rows on first run
                if(allRows.length===0){
                    const initial = Array.from(body.querySelectorAll('tr')).map(r=>r.outerHTML);
                    if(initial.length) allRows = initial;
                }
                const term = document.getElementById('searchBox').value.toLowerCase();
                const pageSize = parseInt(document.getElementById('pageSize').value,10)||20;
                const filtered = allRows.filter(r=>r.toLowerCase().includes(term));
                const total = filtered.length;
                const pages = Math.max(1, Math.ceil(total/pageSize));
                if(currentPage>pages) currentPage=pages;
                const start = (currentPage-1)*pageSize; const end = start+pageSize;
                body.innerHTML = filtered.slice(start,end).join('\n');
                document.getElementById('pageInfo').textContent = `${currentPage}/${pages} (${total})`;
            }
            window.addEventListener('load',()=>{
                // preload rows from server-rendered file (if any)
                const body = document.getElementById('reportBody');
                allRows = Array.from(body.querySelectorAll('tr')).map(r=>r.outerHTML);
                document.getElementById('searchBox').addEventListener('input',()=>{currentPage=1;rebuild();});
                document.getElementById('prevBtn').addEventListener('click',()=>{ if(currentPage>1){currentPage--;rebuild();}});
                document.getElementById('nextBtn').addEventListener('click',()=>{ currentPage++;rebuild();});
                document.getElementById('pageSize').addEventListener('change',()=>{currentPage=1;rebuild();});
                rebuild();
            });
            // expose for other scripts that append rows dynamically
            window.__doccompare = { addRow: function(html){ allRows.unshift(html); rebuild(); }};

            // Auto-refresh toggle (safe default: refresh every 30s)
            let auto = true;
            const refreshInterval = 30*1000;
            setInterval(()=>{ if(auto) rebuild(); }, refreshInterval);
            const toggle = document.createElement('button');
            toggle.textContent = 'AutoRefresh:ON';
            toggle.style.marginLeft = '8px';
            toggle.onclick = ()=>{ auto = !auto; toggle.textContent = 'AutoRefresh:' + (auto? 'ON':'OFF'); };
            document.getElementById('controls').appendChild(toggle);
        })();
    </script>
</body>
</html>
""")
        return index_path


def append_report(report_dir: str, jid: str, aid: str, timestamp: str, artifacts: Dict[str, str], summary: str):
        index_path = _ensure_index(report_dir)
        links_html = []
        # artifacts expected: docx, html, pdf, details
        for label in ("details", "docx", "html", "pdf"):
                if label in artifacts and artifacts[label]:
                        path = artifacts[label]
                        p = path.replace('\\', '/')
                        if label == 'details':
                                display = 'View'
                        else:
                                display = label.upper()
                        links_html.append(f"<a href=\"file:///{p}\" target=\"_blank\">{display}</a>")
        links = " | ".join(links_html)
        row = f"<tr><td>{timestamp}</td><td>{jid}</td><td>{aid}</td><td>{summary}</td><td>{links}</td></tr>\n"
        # append raw row to file so it's available to non-JS clients
        with open(index_path, "r+", encoding="utf-8") as f:
            content = f.read()
            # prefer to prepend newest rows inside <tbody> if present
            tb_open = content.find('<tbody')
            if tb_open != -1:
                # find end of opening tbody tag
                tb_start = content.find('>', tb_open)
                if tb_start != -1:
                    insert_at = tb_start + 1
                    new_content = content[:insert_at] + row + content[insert_at:]
                else:
                    new_content = content + row
            else:
                # fallback: insert a new tbody with the row before </table>
                insert_at = content.rfind('</table>')
                if insert_at != -1:
                    new_content = content[:insert_at] + '<tbody>\n' + row + '</tbody>' + content[insert_at:]
                else:
                    new_content = content + '\n' + row
            f.seek(0); f.truncate(0); f.write(new_content)

        # Also notify client-side script if page is open (optional)
        # This file-based approach cannot push to browsers; browsers must reload to see new rows.


def finalize_index(report_dir: str):
        # nothing to do; index is a complete HTML document created by _ensure_index
        return
