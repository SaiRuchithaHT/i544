import makeApp from './app.js';

const WS_URL = 'https://zdu.binghamton.edu:2345';

window.addEventListener('DOMContentLoaded', async () => {
  await makeApp(WS_URL);
});
