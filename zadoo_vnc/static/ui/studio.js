/* Optional support never gates authentication or a remote-control feature. */
(() => {
  'use strict';
  const dialog = document.getElementById('support-dialog');
  const status = document.getElementById('support-status');
  let returnFocus = null;
  document.querySelectorAll('[data-open-support]').forEach(button => {
    button.addEventListener('click', () => {
      returnFocus = button;
      status.textContent = '';
      const image = document.getElementById('support-qr');
      if (!image.getAttribute('src')) image.src = image.dataset.src;
      dialog.showModal();
      dialog.querySelector('.dialog-close').focus();
    });
  });
  dialog.querySelector('.dialog-close').addEventListener('click', () => dialog.close());
  dialog.addEventListener('click', event => { if (event.target === dialog) {
    const box = dialog.getBoundingClientRect();
    if (event.clientX < box.left || event.clientX > box.right || event.clientY < box.top || event.clientY > box.bottom) dialog.close();
  } });
  dialog.addEventListener('close', () => { if (returnFocus && returnFocus.isConnected) returnFocus.focus({preventScroll:true}); });
  document.getElementById('support-copy-upi').addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText('divyanshgupta675@oksbi');
      status.textContent = 'UPI ID copied. Verify Divyansh Gupta in your payment app.';
    } catch (_) {
      status.textContent = 'Clipboard unavailable. Select and copy the UPI ID shown above.';
    }
  });
})();
