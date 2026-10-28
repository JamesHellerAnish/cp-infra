// Where CPS is and has worked. Coordinates are the town or plant the work is named after,
// not survey points; they place the marker, nothing more.
export type Kind = 'office' | 'hydro' | 'civil' | 'mining';

export interface Site {
  id: string;
  name: string;
  place: string;
  lat: number;
  lon: number;
  kinds: Kind[];
  what: string;
  img?: string;
  href?: string;
  /** Inside the dense Garhwal cluster: selecting it zooms the map in. */
  garhwal?: boolean;
}

export const kindLabel: Record<Kind, string> = {
  office: 'Offices and units',
  hydro: 'Hydropower',
  civil: 'Civil',
  mining: 'Mining',
};

export const sites: Site[] = [
  { id: 'hq', name: 'Head office', place: 'Pitampura, New Delhi', lat: 28.694, lon: 77.153, kinds: ['office'], what: '501, Krishna Apra Business Square, Netaji Subhash Place. Projects are run from here.', href: '/contact/' },
  { id: 'tagena', name: 'Bhutan mine and Tagena unit', place: 'Wangdue, Bhutan', lat: 27.46, lon: 89.9, kinds: ['mining', 'office'], what: '85 acres and 8 million MT of CP-DRUK White Galaxy granite, with a 5-acre processing unit at Tagena.', img: '/media/posters/bhutan-blocks-valley.jpg', href: '/projects/bhutan-mine/' },
  { id: 'rajsamand', name: 'Rajsamand processing unit', place: 'Rajsamand, Rajasthan', lat: 25.07, lon: 73.88, kinds: ['office', 'mining'], what: 'Cuts and finishes the granite quarried at Kekri.', img: '/images/pdf_page_12_img_110.jpg', href: '/projects/kekri-mine/' },
  { id: 'kekri', name: 'Kekri Granite Mine', place: 'Kekri, Rajasthan', lat: 25.97, lon: 75.15, kinds: ['mining'], what: 'One hectare, about 1 million MT: Steel Grey, Black Galaxy and Red Galaxy.', img: '/images/pdf_page_9_img_90.jpg', href: '/projects/kekri-mine/' },

  { id: 'tehri', name: 'Tehri Dam', place: 'Tehri, Uttarakhand', lat: 30.378, lon: 78.48, kinds: ['hydro'], what: 'River dredging, a 47 m baffle wall, HRT/TRT repair, surge shafts and intake epoxy.', img: '/media/posters/tehri-ramp-concreting.jpg', href: '/projects/tehri-dam/', garhwal: true },
  { id: 'koteshwar', name: 'Koteshwar HEP', place: 'Koteshwar, Uttarakhand', lat: 30.262, lon: 78.497, kinds: ['hydro'], what: '36 months of slope stabilisation up to 200 m, then intake and turbine-base grouting.', img: '/images/other_projects/koteshwar_1.jpg', href: '/projects/other/#koteshwar', garhwal: true },
  { id: 'maneri', name: 'Maneri Bhali Stage-I', place: 'Uttarkashi, Uttarakhand', lat: 30.735, lon: 78.53, kinds: ['hydro'], what: 'Spillway and roller bucket repaired under water in the lean season.', img: '/images/other_projects/maneri_bhali_1.jpg', href: '/projects/other/#maneri', garhwal: true },
  { id: 'kedarnath', name: 'Kedarnath Dham', place: 'Kedarnath, Uttarakhand', lat: 30.735, lon: 79.067, kinds: ['civil'], what: 'Pilgrim accommodation blocks Q2Q and Q1N, excavation to finishing.', img: '/images/expansion/kedarnath_q1n_1.jpg', href: '/projects/kedarnath-dham/', garhwal: true },
  { id: 'badrinath', name: 'Badrinath Dham', place: 'Chamoli, Uttarakhand', lat: 30.744, lon: 79.493, kinds: ['hydro', 'civil'], what: 'Riverfronts A and F1, a 47 m footbridge and a civic centre on the Alaknanda.', img: '/media/posters/badrinath-approach.jpg', href: '/projects/badrinath-dham/', garhwal: true },
  { id: 'hailang', name: 'Hailang Dam (VPHEP)', place: 'Helang, Chamoli', lat: 30.523, lon: 79.425, kinds: ['hydro'], what: 'Gate grooves rebuilt in the diversion tunnel with 50–60 cumec still flowing.', img: '/media/posters/hailang-cofferdam-aerial.jpg', href: '/projects/hailang-dam/', garhwal: true },
  { id: 'pipalkoti', name: 'VPHEP Pipalkoti TBM', place: 'Pipalkoti, Chamoli', lat: 30.43, lon: 79.43, kinds: ['hydro'], what: 'A year and a half of PU, micro fine cement and colloidal silica grouting to free a stuck TBM.', img: '/images/other_projects/pipalkoti_grouting_4.jpg', href: '/projects/other/#pipalkoti', garhwal: true },
  { id: 'tapovan', name: 'Tapovan Vishnugad HPP', place: 'Joshimath, Chamoli', lat: 30.495, lon: 79.64, kinds: ['hydro'], what: 'RC dam repair with couplers, shear connectors, micro concrete and epoxy mortar.', img: '/images/other_projects/tapovan_1.jpg', href: '/projects/other/#tapovan', garhwal: true },
  { id: 'sewa', name: 'Sewa II HEP', place: 'Kathua, Jammu & Kashmir', lat: 32.6, lon: 75.78, kinds: ['hydro'], what: 'A 10 km, 3.3 m HRT repaired with epoxy and PU grouting, 2010.', img: '/images/other_projects/sewa_1.jpg', href: '/projects/other/#sewa' },
  { id: 'chamera', name: 'Chamera-III HE Project', place: 'Chamba, Himachal Pradesh', lat: 32.56, lon: 76.24, kinds: ['hydro'], what: 'PU grouting.', href: '/speciality/hydropower/' },

  { id: 'delhi', name: 'Delhi NCR works', place: 'Delhi and Noida', lat: 28.6, lon: 77.3, kinds: ['civil'], what: '12 and 24 MLD CETPs at Nagloi, Bawana and Bhorgarh industrial infrastructure, Sonia Vihar mortar lining, building works in Noida.', img: '/media/posters/noida-building.jpg', href: '/speciality/civil/' },
  { id: 'rohtak', name: 'Industrial Model Township', place: 'Rohtak, Haryana', lat: 28.895, lon: 76.606, kinds: ['civil'], what: 'Township infrastructure.', href: '/speciality/civil/' },
  { id: 'rajpura', name: 'Rajpura Thermal Power', place: 'Rajpura, Punjab', lat: 30.47, lon: 76.6, kinds: ['civil'], what: 'Site grading for 2 × 700 MW, civil works and the Nabha Power railway bridge.', href: '/speciality/civil/' },
  { id: 'bathinda', name: 'HMEL Bathinda Refinery', place: 'Bathinda, Punjab', lat: 30.13, lon: 75.05, kinds: ['civil'], what: 'Specialised insulation.', href: '/speciality/civil/' },
  { id: 'barmer', name: 'Raw water reservoir', place: 'Barmer, Rajasthan', lat: 25.75, lon: 71.39, kinds: ['civil'], what: 'Reservoir excavation.', href: '/speciality/civil/' },
  { id: 'bhopal', name: 'AIIMS Bhopal', place: 'Bhopal, Madhya Pradesh', lat: 23.208, lon: 77.459, kinds: ['civil'], what: 'PU foam insulation and PVC membrane waterproofing.', href: '/speciality/civil/' },
  { id: 'gail', name: 'GAIL gas pipeline', place: 'Jaipur to Loni', lat: 27.83, lon: 76.54, kinds: ['civil'], what: '294 km of pipeline joint coating.', href: '/speciality/civil/' },
  { id: 'expressway', name: 'Nagpur–Mumbai Expressway', place: 'Maharashtra', lat: 20.3, lon: 76.3, kinds: ['civil'], what: 'Passive fire protection, package 14.', href: '/speciality/civil/' },
];

/** Routes drawn as lines: [lat, lon] waypoints. */
export const routes = [
  { id: 'gail', points: [[26.91, 75.79], [27.55, 76.6], [28.75, 77.29]] },
  { id: 'expressway', points: [[21.15, 79.09], [20.7, 77.0], [19.77, 74.48], [19.7, 73.56], [19.22, 73.16]] },
];
