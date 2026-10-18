/* One dispatcher for the module route chain; render only the visible page. */
(() => {
  let frame=0;
  window.addEventListener('hashchange',()=>{cancelAnimationFrame(frame);frame=requestAnimationFrame(()=>route());});
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)window.dispatchEvent(new Event('resize'));});
})();
