// 煒程式工作室：追蹤重要動作（加 LINE、打電話、寄信、填表單、使用試玩）
(function () {
  function send(name, params) { try { if (window.gtag) window.gtag('event', name, params || {}); } catch (e) {} }
  window.wcTrack = send;
  document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a[href]');
    if (!a) return;
    var h = a.getAttribute('href') || '';
    var label = (a.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 40);
    if (/line\.me/.test(h)) send('contact_line', { link_text: label });
    else if (/instagram\.com/.test(h)) send('contact_instagram', { link_text: label });
    else if (/^tel:/.test(h)) send('contact_phone', { link_text: label });
    else if (/^mailto:/.test(h)) send('contact_email', { link_text: label });
    else if (/demo\//.test(h)) send('demo_open', { link_url: h, link_text: label });
  }, true);
  document.addEventListener('submit', function (e) {
    if (e.target && e.target.id === 'inquiry') send('generate_lead', { form: 'inquiry' });
  }, true);
})();
