/* =====================================================================
   SETTINGS: edit this part only
   ===================================================================== */
var CONFIG = {
  name: 'AARF Karaoke & Event Rentals',
  logo: 'images/logo.png',
  // The page asks the Google Sheet "did anything change?" every checkSeconds (5 = your sheet edits show up within a few seconds), and fully reloads every refreshMinutes.
  checkSeconds: 5, refreshMinutes: 5,
  // The whole page also reloads every reloadMinutes, but only when the customer has not started a request.
  reloadMinutes: 10,
  // GCash. Best: paste a fixed-amount payment link (made in PayMongo, HitPay, etc.) into each package's payUrl below.
  // Backup: fill in your GCash number and customers see the exact amount to send.
  gcash: {number:'09997114240', name:''},
  events: ['Birthday','Debut','Reunion','Wedding','Christening','Anniversary','Graduation','Fiesta','Company party','Other'],
  addons: [
    {id:'sm58', title:'Shure SM58 Professional Dynamic Vocal Microphone', note:'', qty:1, price:'150', regular:'200', available:true},
    {id:'tables', title:'Tables', note:'Rental tables for your guests', available:false},
    {id:'chairs', title:'Chairs', note:'Rental chairs for your guests', available:false}
  ],
  // Reminders shown to customers. Delivery: free within freeKm, then perKm for every extra kilometer.
  rules: {deposit:1000, freeKm:5, perKm:30},
  // Amount customers pay on GCash. 0 = the full total rental fee. Put a number (example: 500) to ask only for a reservation fee.
  reservationFee: 0,
  // Paste your Google Apps Script Web App URL here (ends with /exec)
  scriptUrl: 'https://script.google.com/macros/s/AKfycbwjEiND__32BKgkPCXaCefzt1bVSSZWIJQEK2Sw1uYQ8Hd7da_lCChh5DQf41-6BczhAQ/exec',
  // Your Facebook page username, for the optional Messenger button (example: 'myvideokepage'). Leave '' to hide.
  messenger: '1368190183045322',
  // Phone number to show to customers. Leave '' to hide.
  phone: '0999 711 4240',
  // Numbers for the Viber and WhatsApp buttons (same number as your phone by default). Leave '' to hide a button.
  viber: '09997114240',
  whatsapp: '09997114240',
  packages: [
    {id:'standard', icon:'🎤', image:'', payUrl:'', regular:'1,500', title:'Standard Package', price:'1,200', note:'per event, up to 22 hrs',
     items:['Sony ULT Tower 7 420W Speaker','Xiaomi 32 inch LED TV','Platinum XL SD with 23k Songs','Platinum jukebox keypad','Behringer wired mic & 1 Platinum wired mic','Song book 23k songs']},
    {id:'pro', icon:'🔊', image:'', payUrl:'', regular:'1,800', title:'Pro Package', price:'1,500', note:'per event, up to 22 hrs',
     items:['JBL Partybox 710 800W Speaker','Xiaomi 32 inch LED TV','Platinum XL SD with 23k Songs','Platinum jukebox keypad','JBL Partybox Wireless Mic 2','Song book 23k songs','1 High Chair']},
    {id:'premium', icon:'🎉', image:'', payUrl:'', regular:'2,300', title:'Premium Package', price:'1,800', note:'per event, up to 22 hrs',
     items:['JBL Partybox 720 800W Speaker with built-in Battery','Xiaomi 43 inch QLED TV 4k Display','Platinum Piano 4k - 4k Output Display','JBL Easysing Wireless Mic with AI','Song book 23k songs','2 High bar Chair']}
  ]
};
