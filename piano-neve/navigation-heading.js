// GPS course filtering shared by the map and vehicle arrow.
(function () {
 'use strict';
 window.VargaNavigationHeading = function () {
  let heading = null, anchor = null;
  return {
   get value() { return heading; },
   reset() { heading = null; anchor = null; },
   update(coords, ll, distance, bearing) {
    if (!Number.isFinite(coords.accuracy) || coords.accuracy > 35) { anchor = null; return heading; }
    if (Number.isFinite(coords.speed) && coords.speed < 0.8) { anchor = ll.slice(); return heading; }
    let next = null;
    if (Number.isFinite(coords.heading) && coords.heading >= 0 && coords.heading < 360 && Number.isFinite(coords.speed) && coords.speed >= 0.8) next = coords.heading;
    else if (anchor) {
     const moved = distance(anchor, ll), threshold = Math.max(6, Math.min(18, coords.accuracy));
     if (moved >= threshold && moved <= 120) next = bearing(anchor, ll);
     else if (moved > 120) anchor = ll.slice();
    }
    if (!anchor || next !== null) anchor = ll.slice();
    if (next !== null && Number.isFinite(next)) {
     const delta = heading === null ? 0 : ((next - heading + 540) % 360) - 180;
     if (heading === null || Math.abs(delta) >= 65) heading = next;
     else if (Math.abs(delta) >= 3) heading += delta * 0.45;
     heading = (heading + 360) % 360;
    }
    return heading;
   }
  };
 };
})();
