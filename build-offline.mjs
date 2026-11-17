// 오프라인용 단일 HTML 파일 만들기: node build-offline.mjs
import fs from 'node:fs';
const read = f => fs.readFileSync(new URL(f, import.meta.url), 'utf8');
const js = s => s.replace(/<\/script/gi, '<\\/script');
let html = read('./index.html');
html = html.replace(/script-src 'self'/, "script-src 'unsafe-inline'").replace(/worker-src 'self'/, "worker-src 'none'");
html = html.replace(/<link rel="stylesheet" href="app\.css[^"]*">/, () => '<style>\n' + read('./app.css') + '\n</style>');
html = html.replace(/<a class="btn ghost small" id="offlineLink"[^>]*>[^<]*<\/a>/, '');
html = html.replace(/<script src="vendor\/xlsx\.bundle\.js[^"]*"><\/script>/, () => '<script>\n' + js(read('./vendor/xlsx.bundle.js')) + '\n</script>');
html = html.replace(/<script src="engine\.js[^"]*"><\/script>/, () => '<script>\n' + js(read('./engine.js')) + '\n</script>');
html = html.replace(/<script src="app\.js[^"]*"><\/script>/, () => '<script>\nwindow.__OFFLINE__=true;\n' + js(read('./app.js')) + '\n</script>');
html = html.replace('<title>신입생 반편성 도우미</title>', '<title>신입생 반편성 도우미 (오프라인)</title>');
fs.mkdirSync(new URL('./offline/', import.meta.url), { recursive: true });
fs.writeFileSync(new URL('./offline/class-assign-offline.html', import.meta.url), html);
console.log('offline/class-assign-offline.html', (html.length / 1024).toFixed(0) + 'KB', 'external refs:', (html.match(/(src|href)="(?!#|data:)[^"]+"/g) || []).join(' '));

