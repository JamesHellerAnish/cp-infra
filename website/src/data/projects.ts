// The project index shared by the home page, the menus and the cross-links at the foot of each project.
export interface Project {
  name: string;
  href: string;
  what: string;
  place: string;
}

export const projects: Project[] = [
  { name: 'Tehri Dam', href: '/projects/tehri-dam/', what: 'River dredging, baffle wall, HRT/TRT and intake repair', place: 'Tehri, Uttarakhand' },
  { name: 'Badrinath Dham', href: '/projects/badrinath-dham/', what: 'Three riverfronts, a 47 m footbridge and a civic centre', place: 'Chamoli, Uttarakhand' },
  { name: 'Hailang Dam (VPHEP)', href: '/projects/hailang-dam/', what: 'Gate groove restoration in the diversion tunnel', place: 'Helang, Uttarakhand' },
  { name: 'Kedarnath Dham', href: '/projects/kedarnath-dham/', what: 'Pilgrim accommodation blocks Q2Q and Q1N', place: 'Kedarnath, Uttarakhand' },
  { name: 'Bhutan Granite Mine', href: '/projects/bhutan-mine/', what: '85 acres of CP-DRUK White Galaxy granite', place: 'Wangdue, Bhutan' },
  { name: 'Kekri Granite Mine', href: '/projects/kekri-mine/', what: 'Steel Grey, Black Galaxy and Red Galaxy granite', place: 'Kekri, Rajasthan' },
  { name: 'Other projects', href: '/projects/other/', what: 'Koteshwar, Maneri Bhali, Tapovan, Pipalkoti and more', place: 'Across India' },
];

/** The other projects, starting after the current one, for the "next" index at the foot of a project page. */
export function others(href: string): Project[] {
  const i = projects.findIndex((p) => p.href === href);
  return [...projects.slice(i + 1), ...projects.slice(0, Math.max(i, 0))];
}
