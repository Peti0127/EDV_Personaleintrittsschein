/* Lagepläne per Standort (Bilder in Images/Lageplaene/).
   Ein Plan:              "Standort": { code: "ATSxx", src: "Images/Lageplaene/..." }
   Ein Plan pro Stock:    "Standort": { code: "ATSxx", floors: { "Erdgeschoß": "...", "1. Stock": "..." } }
   Pro Gebäude und Stock: "Standort": { code: "ATSxx", buildings: { "Gebäude": { "Stock": "..." } } } */
const LAGEPLAENE = {
  "Klosterneuburg": {
    "code": "ATS01",
    "buildings": {
      "Hauptgebäude": {
        "Erdgeschoß": "Images/Lageplaene/ATS01_HG_EG_Lageplan.png",
        "1. Stock": "Images/Lageplaene/ATS01_HG_1_Stock_Lageplan.png",
        "2. Stock": "Images/Lageplaene/ATS01_HG_2_Stock_Lageplan.png"
      },
      "Neubau": {
        "Erdgeschoß": "Images/Lageplaene/ATS01_NB_EG_Lageplan.png",
        "1. Stock": "Images/Lageplaene/ATS01_NB_1_Stock_Lageplan.png"
      },
      "Inku": {
        "Erdgeschoß Links": "Images/Lageplaene/ATS01_IN_EG_Links_Lageplan.png",
        "Obergeschoß Links": "Images/Lageplaene/ATS01_IN_OG_Links_Lageplan.png",
        "Obergeschoß Rechts": "Images/Lageplaene/ATS01_IN_OG_Rechts_Lageplan.png"
      }
    }
  },
  "Kalsdorf": {
    "code": "ATS02",
    "src": "Images/Lageplaene/ATS02_Lageplan.jpg"
  },
  "Klagenfurt": {
    "code": "ATS05",
    "src": "Images/Lageplaene/ATS05_Lageplan.png"
  },
  "Leoben": {
    "code": "ATS06",
    "src": "Images/Lageplaene/ATS06_Lageplan.png"
  },
  "Salzburg": {
    "code": "ATS09",
    "src": "Images/Lageplaene/ATS09_Lageplan.png"
  },
  "Oberwart": {
    "code": "ATS14",
    "src": "Images/Lageplaene/ATS14_Lageplan.png"
  },
  "Innsbruck": {
    "code": "ATS03",
    "src": "Images/Lageplaene/ATS03_Lageplan.png"
  },
  "Pinzgau": {
    "code": "ATS08",
    "src": "Images/Lageplaene/ATS08_Lageplan.png"
  },
  "Spillern": {
    "code": "ATS11",
    "src": "Images/Lageplaene/ATS11_Lageplan.png"
  },
  "Rankweil": {
    "code": "ATS16",
    "src": "Images/Lageplaene/ATS16_Lageplan.png"
  },
  "Krems": {
    "code": "ATS23",
    "src": "Images/Lageplaene/ATS23_Lageplan.png"
  },
  "Neusiedl": {
    "code": "ATS32",
    "src": "Images/Lageplaene/ATS32_Lageplan.png"
  },
  "St. Pölten": {
    "code": "ATS12",
    "src": "Images/Lageplaene/ATS12_Lageplan.png"
  },
  "Linz": {
    "code": "ATS24",
    "src": "Images/Lageplaene/ATS24_Lageplan.png"
  },
  "Wr. Neustadt": {
    "code": "ATS18",
    "floors": {
      "Erdgeschoß": "Images/Lageplaene/ATS18_EG_Lageplan.png",
      "1. Stock": "Images/Lageplaene/ATS18_OG_Lageplan.png"
    }
  }
};
