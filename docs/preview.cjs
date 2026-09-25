/* Local, mock-only design preview. Never connects to Google or stores keys.
   Run npm run preview; this server is NOT the Apps Script deployment. */
const http=require('node:http');
const fs=require('node:fs');
const path=require('node:path');
const root=path.resolve(__dirname,'..');
const source=fs.readFileSync(path.join(root,'WebApp.html'),'utf8');
const html=source.replace('  <script>\n  (function () {','  <script src="/preview-mock.cjs"></script>\n  <script>\n  (function () {');
const port=Number(process.env.PORT)||4173;
http.createServer((req,res)=>{
  if(req.url==='/' || req.url==='/index.html'){
    res.writeHead(200,{'content-type':'text/html; charset=utf-8','cache-control':'no-store'});res.end(html);return;
  }
  if(req.url==='/preview-mock.cjs'){
    res.writeHead(200,{'content-type':'application/javascript; charset=utf-8','cache-control':'no-store'});
    res.end(fs.readFileSync(path.join(__dirname,'preview-mock.cjs')));return;
  }
  res.writeHead(404,{'content-type':'text/plain; charset=utf-8'});res.end('Not found.');
}).listen(port,'0.0.0.0',()=>console.log('Lesson Grader design preview (mock data only) listening on port '+port));
