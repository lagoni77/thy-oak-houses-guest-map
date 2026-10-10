// Stand-in for Code.gs parseAirbnbEmail in unit tests (synthetic).
function parseAirbnbEmail(body) {
  if (/BROKEN/.test(body)) return null;
  if (/THROW/.test(body)) throw new Error('boom');
  return { listing: 'Tiny Oak House', name: 'Synthetic', checkIn: '1999-01-01', checkOut: '1999-01-02', nights: 1,
           bookingUrl: 'https://www.airbnb.dk/hosting/reservations/details/HMAAAAAAA1' };
}
