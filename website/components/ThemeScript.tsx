// Inlined in <head> before hydration to prevent theme flash.
export default function ThemeScript() {
  const code = `(function(){try{var t=localStorage.getItem('theme');if(!t){t='dark';}document.documentElement.classList.remove('light','dark');document.documentElement.classList.add(t);}catch(e){document.documentElement.classList.add('dark');}})();`;
  return <script dangerouslySetInnerHTML={{ __html: code }} />;
}
