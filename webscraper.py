import excel

import time

from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys

# Because the postal codes of the belgian cities are static, I made a program that once searched all the city names of
# the associated postal codes and copied the outcome into this dictionary.
dictionary_belgian_postal_codes_with_city_names = {1000: 'Brussel', 1070: 'Anderlecht', 1050: 'Elsene',
                                                   1040: 'Etterbeek', 1140: 'Evere', 1083: 'Ganshoren', 1130: 'Haren',
                                                   1090: 'Jette', 1081: 'Koekelberg', 1020: 'Laken',
                                                   1120: 'Neder-Over-Heembeek', 1160: 'Oudergem', 1030: 'Schaarbeek',
                                                   1082: 'Sint-Agatha-Berchem', 1060: 'Sint-Gillis',
                                                   1080: 'Sint-Jans-Molenbeek', 1210: 'Sint-Joost-Ten-Node',
                                                   1200: 'Sint-Lambrechts-Woluwe', 1150: 'Sint-Pieters-Woluwe',
                                                   1180: 'Ukkel', 1190: 'Vorst', 1170: 'Watermaal-Bosvoorde',
                                                   2000: 'Antwerpen', 2018: 'Antwerpen', 2020: 'Antwerpen',
                                                   2030: 'Antwerpen', 2040: 'Antwerpen', 2050: 'Antwerpen',
                                                   2060: 'Antwerpen', 2070: 'Burcht', 2100: 'Deurne', 2110: 'Wijnegem',
                                                   2140: 'Borgerhout', 2150: 'Borsbeek', 2160: 'Wommelgem',
                                                   2170: 'Merksem', 2180: 'Ekeren', 2200: 'Herentals',
                                                   2220: 'Heist-op-den-Berg', 2221: 'Booischot', 2222: 'Wiekevorst',
                                                   2223: 'Schriek', 2230: 'Ramsel', 2235: 'Westmeerbeek',
                                                   2240: 'Massenhoven', 2242: 'Pulderbos', 2243: 'Pulle', 2250: 'Olen',
                                                   2260: 'Zoerle-Parwijs', 2270: 'Herenthout', 2275: 'Poederlee',
                                                   2280: 'Grobbendonk', 2288: 'Bouwel', 2290: 'Vorselaar',
                                                   2300: 'Turnhout', 2310: 'Rijkevorsel', 2320: 'Hoogstraten',
                                                   2321: 'Meer', 2322: 'Minderhout', 2323: 'Wortel', 2328: 'Meerle',
                                                   2330: 'Merksplas', 2340: 'Beerse', 2350: 'Vosselaar',
                                                   2360: 'Oud-Turnhout', 2370: 'Arendonk', 2380: 'Ravels',
                                                   2381: 'Weelde', 2382: 'Poppel', 2387: 'Baarle-Hertog',
                                                   2390: 'Westmalle', 2400: 'Mol', 2430: 'Vorst', 2431: 'Veerle',
                                                   2440: 'Geel', 2450: 'Meerhout', 2460: 'Lichtaart', 2470: 'Retie',
                                                   2480: 'Dessel', 2490: 'Balen', 2491: 'Olmen', 2500: 'Lier',
                                                   2520: 'Broechem', 2530: 'Boechout', 2531: 'Vremde', 2540: 'Hove',
                                                   2547: 'Lint', 2550: 'Kontich', 2560: 'Nijlen', 2570: 'Duffel',
                                                   2580: 'Beerzel', 2590: 'Berlaar', 2600: 'Berchem', 2610: 'Wilrijk',
                                                   2620: 'Hemiksem', 2627: 'Schelle', 2630: 'Aartselaar',
                                                   2640: 'Mortsel', 2650: 'Edegem', 2660: 'Hoboken', 2800: 'Mechelen',
                                                   2801: 'Heffen', 2811: 'Leest', 2812: 'Muizen', 2820: 'Rijmenam',
                                                   2830: 'Blaasveld', 2840: 'Reet', 2845: 'Niel', 2850: 'Boom',
                                                   2860: 'Sint-Katelijne-Waver', 2861: 'Onze-Lieve-Vrouw-Waver',
                                                   2870: 'Liezele', 2880: 'Weert', 2890: 'Oppuurs', 2900: 'Schoten',
                                                   2910: 'Essen', 2920: 'Kalmthout', 2930: 'Brasschaat',
                                                   2940: 'Hoevenen', 2950: 'Kapellen', 2960: 'Sint-Lenaarts',
                                                   2970: 'Schilde', 2980: 'Halle', 2990: 'Wuustwezel', 9031: 'Drongen',
                                                   6000: 'Charleroi', 6001: 'Marcinelle', 6010: 'Couillet',
                                                   6020: 'Dampremy', 6030: 'Marchienne-au-pont',
                                                   6031: 'Monceau-sur-sambre', 6032: 'Mont-sur-marchienne',
                                                   6040: 'Jumet', 6041: 'Gosselies', 6042: 'Lodelinsart',
                                                   6043: 'Ransart', 6044: 'Roux', 6060: 'Gilly',
                                                   6061: 'Montignies-sur-sambre', 6110: 'Montigny-le-tilleul',
                                                   6111: 'Landelies', 6120: 'Cour-sur-heure', 6140: "Fontaine-l'evêque",
                                                   6141: 'Forchies-la-marche', 6142: 'Leernes', 6150: 'Anderlues',
                                                   6180: 'Courcelles', 6181: 'Gouy-lez-piéton', 6182: 'Souvret',
                                                   6183: 'Trazegnies', 6200: 'Châtelet', 6210: 'Rèves', 6211: 'Mellet',
                                                   6220: 'Fleurus', 6221: 'Saint-amand', 6222: 'Brye', 6223: 'Wagnelée',
                                                   6224: 'Wanfercée-baulet', 6230: 'Buzet', 6238: 'Liberchies',
                                                   6240: 'Farciennes', 6250: 'Roselies', 6280: 'Acoz',
                                                   6440: 'Fourbechies', 6441: 'Erpion', 6460: 'Robechies',
                                                   6461: 'Virelles', 6462: 'Vaulx-lez-chimay', 6463: 'Lompret',
                                                   6464: 'Rièzes', 6470: 'Rance', 6500: 'Thirimont', 6511: 'Strée',
                                                   6530: 'Leers-et-fosteau', 6531: 'Biesme-sous-thuin', 6532: 'Ragnies',
                                                   6533: 'Biercée', 6534: 'Gozée', 6536: 'Donstiennes', 6540: 'Lobbes',
                                                   6542: 'Sars-la-buissière', 6543: 'Bienne-lez-happart',
                                                   6560: "Bersillies-l'abbaye", 6567: 'Labuissière', 6590: 'Momignies',
                                                   6591: 'Macon', 6592: 'Monceau-imbrechies', 6593: 'Macquenoise',
                                                   6594: 'Beauwelz', 6596: 'Seloignes', 7000: 'Mons', 7011: 'Ghlin',
                                                   7012: 'Flénu', 7020: 'Nimy', 7021: 'Havre', 7022: 'Mesvin',
                                                   7024: 'Ciply', 7030: 'Saint-symphorien',
                                                   7031: 'Villers-saint-ghislain', 7032: 'Spiennes', 7033: 'Cuesmes',
                                                   7034: 'Saint-denis', 7040: 'Aulnois', 7041: 'Havay', 7050: 'Jurbise',
                                                   7060: 'Horrues', 7061: 'Thieusies', 7062: 'Naast',
                                                   7063: 'Neufvilles', 7070: 'Mignault', 7080: 'Noirchain',
                                                   7090: 'Steenkerque', 7100: 'La louvière', 7110: 'Houdeng-goegnies',
                                                   7120: 'Estinnes-au-val', 7130: 'Battignies', 7131: 'Waudrez',
                                                   7133: 'Buvrinnes', 7134: 'Leval-trahegnies',
                                                   7140: 'Morlanwelz-mariemont', 7141: 'Mont-sainte-aldegonde',
                                                   7160: 'Chapelle-lez-herlaimont', 7170: "Bois-d'haine",
                                                   7180: 'Seneffe', 7181: 'Petit-roeulx-lez-nivelles',
                                                   7190: "Ecaussinnes-d'enghien", 7191: 'Ecaussinnes-lalaing',
                                                   7300: 'Boussu', 7301: 'Hornu', 7320: 'Bernissart', 7321: 'Harchies',
                                                   7322: 'Ville-pommeroeul', 7330: 'Saint-ghislain', 7331: 'Baudour',
                                                   7332: 'Neufmaison', 7333: 'Tertre', 7334: 'Hautrage', 7340: 'Wasmes',
                                                   7350: 'Montroeul-sur-haine', 7370: 'Dour', 7380: 'Quiévrain',
                                                   7382: 'Audregnies', 7387: 'Autreppe', 7390: 'Quaregnon',
                                                   7500: 'Saint-maur', 7501: 'Orcq', 7502: 'Esplechin',
                                                   7503: 'Froyennes', 7504: 'Froidmont', 7506: 'Willemeau',
                                                   7520: 'Ramegnies-chin', 7521: 'Chercq', 7522: 'Marquain',
                                                   7530: 'Gaurain-ramecroix', 7531: 'Havinnes', 7532: 'Beclers',
                                                   7533: 'Thimougies', 7534: 'Barry', 7536: 'Vaulx', 7538: 'Vezon',
                                                   7540: 'Rumillies', 7542: 'Mont-saint-aubert', 7543: 'Mourcourt',
                                                   7548: 'Warchin', 7600: 'Péruwelz', 7601: 'Roucourt', 7602: 'Bury',
                                                   7603: 'Bon-secours', 7604: 'Brasmenil', 7608: 'Wiers', 7610: 'Rumes',
                                                   7611: 'La glanerie', 7618: 'Taintignies', 7620: 'Guignies',
                                                   7621: 'Lesdain', 7622: 'Laplaigne', 7623: 'Rongy',
                                                   7624: 'Howardries', 7640: 'Antoing', 7641: 'Bruyelle',
                                                   7642: 'Calonne', 7643: 'Fontenoy', 7700: 'Luingne',
                                                   7711: 'Dottenijs', 7712: 'Herseaux', 7730: 'Saint-léger',
                                                   7740: 'Warcoing', 7742: 'Hérinnes-lez-pecq', 7743: 'Esquelmes',
                                                   7750: "Mont-de-l'enclus", 7760: 'Pottes', 7780: 'Komen-waasten',
                                                   7781: 'Houthem', 7782: 'Ploegsteert', 7783: 'Bizet', 7784: 'Waasten',
                                                   7800: 'Lanquesaint', 7801: 'Irchonwelz', 7802: 'Ormeignies',
                                                   7803: 'Bouvignies', 7804: 'Rebaix', 7810: 'Maffle', 7811: 'Arbre',
                                                   7812: 'Mainvault', 7822: 'Ghislenghien', 7823: 'Gibecq',
                                                   7830: 'Hoves', 7850: 'Mark', 7860: 'Lessines', 7861: 'Wannebecq',
                                                   7862: 'Ogy', 7863: 'Ghoy', 7864: 'Deux-acren',
                                                   7866: 'Bois-de-lessines', 7870: 'Bauffe', 7880: 'Vloesberg',
                                                   7890: 'Wodecq', 7900: 'Leuze-en-hainaut', 7901: 'Thieulain',
                                                   7903: 'Blicquy', 7904: 'Tourpes', 7906: 'Gallaix',
                                                   7910: 'Frasnes-lez-anvaing', 7911: 'Hacquegnies', 7912: 'Dergneau',
                                                   7940: 'Brugelette', 7941: 'Attre', 7942: 'Mévergnies-lez-lens',
                                                   7943: 'Gages', 7950: 'Ladeuze', 7951: 'Tongre-notre-dame',
                                                   7970: 'Beloeil', 7971: 'Wadelincourt', 7972: 'Aubechies',
                                                   7973: 'Stambruges', 3500: 'Sint-lambrechts-herk',
                                                   3501: 'Wimmertingen', 3510: 'Kermt', 3511: 'Stokrooie',
                                                   3512: 'Stevoort', 3520: 'Zonhoven', 3530: 'Helchteren',
                                                   3540: 'Schulen', 3545: 'Loksbergen', 3550: 'Zolder', 3560: 'Meldert',
                                                   3570: 'Alken', 3580: 'Beringen', 3581: 'Beverlo', 3582: 'Koersel',
                                                   3583: 'Paal', 3590: 'Diepenbeek', 3600: 'Genk', 3620: 'Lanaken',
                                                   3621: 'Rekem', 3630: 'Opgrimbie', 3631: 'Uikhoven', 3640: 'Kinrooi',
                                                   3650: 'Elen', 3660: 'Oudsbergen', 3665: 'As', 3668: 'Niel-bij-As',
                                                   3670: 'Neerglabbeek', 3680: 'Opoeteren', 3690: 'Zutendaal',
                                                   3700: 'Haren', 3717: 'Herstappe', 3720: 'Kortessem',
                                                   3721: 'Vliermaalroot', 3722: 'Wintershoven', 3723: 'Guigoven',
                                                   3724: 'Vliermaal', 3730: 'Hoeselt', 3732: 'Schalkhoven',
                                                   3740: 'Rosmeer', 3742: 'Martenslinde', 3746: 'Hoelbeek',
                                                   3770: 'Herderen', 3790: 'Sint-Martens-Voeren', 3791: 'Remersdaal',
                                                   3792: 'Sint-Pieters-Voeren', 3793: 'Teuven', 3798: "'s-Gravenvoeren",
                                                   3800: 'Gelinden', 3803: 'Gorsem', 3806: 'Velm', 3830: 'Wellen',
                                                   3831: 'Herten', 3832: 'Ulbeek', 3840: 'Hendrieken', 3850: 'Kozen',
                                                   3870: 'Rukkelingen-Loon', 3890: 'Vorsen', 3891: 'Buvingen',
                                                   3900: 'Overpelt', 3910: 'Sint-Huibrechts-Lille', 3920: 'Lommel',
                                                   3930: 'Hamont-Achel', 3940: 'Hechtel-Eksel', 3941: 'Eksel',
                                                   3945: 'Oostham', 3950: 'Bocholt', 3960: 'Beek', 3970: 'Leopoldsburg',
                                                   3971: 'Heppen', 3980: 'Tessenderlo', 3990: 'Wijchmaal', 4000: 'Luik',
                                                   4020: 'Liège', 4030: 'Grivegnee', 4031: 'Angleur', 4032: 'Chênee',
                                                   4040: 'Herstal', 4041: 'Vottem', 4042: 'Liers',
                                                   4050: 'Chaudfontaine', 4051: 'Vaux-sous-chèvremont',
                                                   4052: 'Beaufays', 4053: 'Embourg', 4100: 'Boncelles',
                                                   4101: 'Jemeppe-sur-meuse', 4102: 'Ougrée', 4120: 'Ehein',
                                                   4121: 'Neuville-en-condroz', 4122: 'Plainevaux', 4130: 'Esneux',
                                                   4140: 'Gomzé-andoumont', 4141: 'Louveigné', 4160: 'Anthisnes',
                                                   4161: 'Villers-aux-tours', 4162: 'Hody', 4163: 'Tavier',
                                                   4170: 'Comblain-au-pont', 4171: 'Poulseur', 4180: 'Hamoir',
                                                   4181: 'Filot', 4190: 'Werbomont', 4210: 'Hannêche', 4217: 'Héron',
                                                   4218: 'Couthuin', 4219: 'Ambresin', 4250: 'Boëlhe', 4252: 'Omal',
                                                   4253: 'Darion', 4254: 'Ligney', 4257: 'Rosoux-crenwick',
                                                   4260: 'Avennes', 4261: 'Latinne', 4263: 'Tourinne', 4280: 'Hannut',
                                                   4287: 'Racour', 4300: 'Lantremange', 4317: 'Viemme', 4340: 'Othée',
                                                   4342: 'Hognoul', 4347: 'Noville', 4350: 'Lamine', 4351: 'Hodeige',
                                                   4357: 'Donceel', 4360: 'Otrange', 4367: 'Fize-le-marsal',
                                                   4400: 'Mons-lez-liège', 4420: 'Tilleur', 4430: 'Ans', 4431: 'Loncin',
                                                   4432: 'Alleur', 4450: 'Juprelle', 4451: 'Voroux-lez-liers',
                                                   4452: 'Wihogne', 4453: 'Villers-saint-siméon', 4458: 'Fexhe-slins',
                                                   4460: 'Grâce-berleur', 4470: 'Saint-georges-sur-meuse',
                                                   4480: 'Hermalle-sous-huy', 4500: 'Tihange', 4520: 'Wanze',
                                                   4530: 'Vieux-waleffe', 4537: 'Bodegnée', 4540: 'Amay',
                                                   4550: 'Saint-séverin', 4557: 'Fraiture', 4560: 'Ocquier',
                                                   4570: 'Marchin', 4577: 'Vierset-barse', 4590: 'Ouffet', 4600: 'Visé',
                                                   4601: 'Argenteau', 4602: 'Cheratte', 4606: 'Saint-andré',
                                                   4607: 'Berneau', 4608: 'Neufchâteau', 4610: 'Queue-du-bois',
                                                   4620: 'Fléron', 4621: 'Retinne', 4623: 'Magnée', 4624: 'Romsée',
                                                   4630: 'Soumagne', 4631: 'Evegnée', 4632: 'Cérexhe-heuseux',
                                                   4633: 'Melen', 4650: 'Herve', 4651: 'Battice', 4652: 'Xhendelesse',
                                                   4653: 'Bolland', 4654: 'Charneux', 4670: 'Trembleur', 4671: 'Housse',
                                                   4672: 'Saint-remy', 4680: 'Hermée', 4681: 'Hermalle-sous-argenteau',
                                                   4682: 'Heure-le-romain', 4683: 'Vivegnis', 4684: 'Haccourt',
                                                   4690: 'Bassenge', 4700: 'Eupen', 4701: 'Kettenis', 4710: 'Lontzen',
                                                   4711: 'Walhorn', 4720: 'Kelmis', 4721: 'Neu-moresnet',
                                                   4728: 'Hergenrath', 4730: 'Raeren', 4731: 'Eynatten',
                                                   4750: 'Butgenbach', 4760: 'Manderfeld', 4761: 'Rocherath',
                                                   4770: 'Meyerode', 4771: 'Heppenbach', 4780: 'Recht',
                                                   4782: 'Schönberg', 4783: 'Lommersweiler', 4784: 'Crombach',
                                                   4790: 'Burg-reuland', 4791: 'Thommen', 4800: 'Petit-rechain',
                                                   4801: 'Stembert', 4802: 'Heusy', 4820: 'Dison', 4821: 'Andrimont',
                                                   4830: 'Limbourg', 4831: 'Bilstain', 4834: 'Goé', 4837: 'Membach',
                                                   4840: 'Welkenraedt', 4841: 'Henri-chapelle', 4845: 'Sart-lez-spa',
                                                   4850: 'Plombières', 4851: 'Sippenaeken', 4852: 'Hombourg',
                                                   4860: 'Cornesse', 4861: 'Soiron', 4870: 'Fraipont', 4877: 'Olne',
                                                   4880: 'Aubel', 4890: 'Thimister-clermont', 4900: 'Spa',
                                                   4910: 'Theux', 4920: 'Sougné-remouchamps', 4950: 'Sourbrodt',
                                                   4960: 'Bellevaux-ligneuville', 4970: 'Francorchamps', 4980: 'Fosse',
                                                   4983: 'Basse-bodeux', 4987: 'Chevron', 4990: 'Arbrefontaine',
                                                   6600: 'Longvilly', 6630: 'Martelange', 6637: 'Tintange',
                                                   6640: 'Vaux-sur-sûre', 6642: 'Juseret', 6660: 'Nadrin', 6661: 'Mont',
                                                   6662: 'Tavigny', 6663: 'Mabompré', 6666: 'Wibrin', 6670: 'Gouvy',
                                                   6671: 'Bovigny', 6672: 'Beho', 6673: 'Cherain', 6674: 'Montleban',
                                                   6680: 'Sainte-ode', 6681: 'Lavacherie', 6686: 'Flamierge',
                                                   6687: 'Bertogne', 6688: 'Longchamps', 6690: 'Bihain',
                                                   6692: 'Petit-thier', 6698: 'Grand-halleux', 6700: 'Toernich',
                                                   6704: 'Guirsch', 6706: 'Autelbas', 6717: 'Nothomb', 6720: 'Hachy',
                                                   6721: 'Anlier', 6723: 'Habay-la-vieille', 6724: 'Rulles',
                                                   6730: 'Tintigny', 6740: 'Sainte-marie-sur-semois', 6741: 'Vance',
                                                   6742: 'Chantemelle', 6743: 'Buzenol', 6747: 'Meix-le-tige',
                                                   6750: 'Mussy-la-ville', 6760: 'Ruette', 6761: 'Latour',
                                                   6762: 'Saint-mard', 6767: 'Rouvroy', 6769: 'Sommethonne',
                                                   6780: 'Wolkrange', 6781: 'Sélange', 6782: 'Habergy', 6790: 'Aubange',
                                                   6791: 'Athus', 6792: 'Rachecourt', 6800: 'Recogne', 6810: 'Izel',
                                                   6811: 'Les bulles', 6812: 'Suxy', 6813: 'Termes',
                                                   6820: 'Fontenoille', 6821: 'Lacuisine', 6823: 'Villers-devant-orval',
                                                   6824: 'Chassepierre', 6830: 'Poupehan', 6831: 'Noirfontaine',
                                                   6832: 'Sensenruth', 6833: 'Vivy', 6834: 'Bellevaux', 6836: 'Dohan',
                                                   6838: 'Corbion', 6840: 'Grapfontaine', 6850: 'Offagne',
                                                   6851: 'Nollevaux', 6852: 'Maissin', 6853: 'Framont',
                                                   6856: 'Fays-les-veneurs', 6860: 'Ebly', 6870: 'Vesqueville',
                                                   6880: 'Bertrix', 6887: 'Straimont', 6890: 'Ochamps', 6900: 'Waha',
                                                   6920: 'Sohier', 6921: 'Chanly', 6922: 'Halma', 6924: 'Lomprez',
                                                   6927: 'Bure', 6929: 'Haut-fays', 6940: 'Durbuy', 6941: 'Heyd',
                                                   6950: 'Harsin', 6951: 'Bande', 6952: 'Grune', 6953: 'Forrières',
                                                   6960: 'Odeigne', 6970: 'Tenneville', 6971: 'Champlon',
                                                   6972: 'Erneuville', 6980: 'Beausaint', 6982: 'Samrée', 6983: 'Ortho',
                                                   6984: 'Hives', 6986: 'Halleux', 6987: 'Beffe', 6990: 'Hotton',
                                                   6997: 'Soy', 5000: 'Beez', 5001: 'Belgrade', 5002: 'Saint-servais',
                                                   5003: 'Saint-marc', 5004: 'Bouge', 5020: 'Suarlée', 5021: 'Boninne',
                                                   5022: 'Cognelée', 5024: 'Gelbressée', 5030: 'Gembloux',
                                                   5031: 'Grand-leez', 5032: 'Mazy', 5060: 'Sambreville',
                                                   5070: 'Aisemont', 5080: 'La bruyère', 5081: 'Meux', 5100: 'Wierde',
                                                   5101: 'Erpent', 5140: 'Boignée', 5150: 'Soye', 5170: 'Lesve',
                                                   5190: 'Spy', 5300: 'Landenne', 5310: 'Noville-sur-méhaigne',
                                                   5330: 'Maillen', 5332: 'Crupet', 5333: 'Sorinne-la-longue',
                                                   5334: 'Florée', 5336: 'Courrière', 5340: 'Sorée', 5350: 'Evelette',
                                                   5351: 'Haillot', 5352: 'Perwez-haillot', 5353: 'Goesnes',
                                                   5354: 'Jallet', 5360: 'Hamois', 5361: 'Scy', 5362: 'Achet',
                                                   5363: 'Emptinne', 5364: 'Schaltin', 5370: 'Porcheresse',
                                                   5372: 'Méan', 5374: 'Maffe', 5376: 'Miécret', 5377: 'Heure',
                                                   5380: 'Noville-les-bois', 5500: 'Dréhance', 5501: 'Lisogne',
                                                   5502: 'Thynes', 5503: 'Sorinnes', 5504: 'Foy-notre-dame',
                                                   5520: 'Anthée', 5521: 'Serville', 5522: 'Falaen', 5523: 'Sommière',
                                                   5524: 'Gerin', 5530: 'Evrehailles', 5537: 'Warnant',
                                                   5540: 'Hermeton-sur-meuse', 5541: 'Hastière-par-delà',
                                                   5542: 'Blaimont', 5543: 'Heer', 5544: 'Agimont', 5550: 'Bohan',
                                                   5555: 'Monceau-en-ardenne', 5560: 'Finnevaux', 5561: 'Celles',
                                                   5562: 'Custinne', 5563: 'Hour', 5564: 'Wanlin', 5570: 'Vonêche',
                                                   5571: 'Wiesme', 5572: 'Focant', 5573: 'Martouzin-neuville',
                                                   5574: 'Pondrôme', 5575: 'Bourseigne-neuve', 5576: 'Froidfontaine',
                                                   5580: 'Wavreille', 5590: 'Conneux', 5600: 'Fagnolle',
                                                   5620: 'Corenne', 5621: 'Hanzinne', 5630: 'Soumoy', 5640: 'Oret',
                                                   5641: 'Furnaux', 5644: 'Ermeton-sur-biert', 5646: 'Stave',
                                                   5650: 'Vogenée', 5651: 'Somzée', 5660: 'Presgaux',
                                                   5670: 'Vierves-sur-viroin', 5680: 'Vodelée', 8550: 'Zwevegem',
                                                   9000: 'Gent', 9030: 'Mariakerke', 9031: 'Drongen', 9032: 'Wondelgem',
                                                   9040: 'Sint-amandsberg', 9041: 'Oostakker',
                                                   9042: 'Sint-kruis-winkel', 9050: 'Ledeberg', 9051: 'Afsnee',
                                                   9052: 'Zwijnaarde', 9060: 'Zelzate', 9070: 'Destelbergen',
                                                   9080: 'Zaffelare', 9090: 'Gontrode', 9100: 'Nieuwkerken-Waas',
                                                   9111: 'Belsele', 9112: 'Sinaai-Waas', 9120: 'Haasdonk', 9130: 'Doel',
                                                   9140: 'Tielrode', 9150: 'Rupelmonde', 9160: 'Daknam',
                                                   9170: 'Sint-Gillis-Waas', 9180: 'Moerbeke-waas', 9185: 'Wachtebeke',
                                                   9190: 'Kemzeke', 9200: 'Baasrode', 9220: 'Hamme', 9230: 'Massemen',
                                                   9240: 'Zele', 9250: 'Waasmunster', 9255: 'Buggenhout',
                                                   9260: 'Wichelen', 9270: 'Laarne', 9280: 'Lebbeke', 9290: 'Uitbergen',
                                                   9300: 'Aalst', 9308: 'Gijzegem', 9310: 'Moorsel',
                                                   9320: 'Nieuwerkerken', 9340: 'Impe', 9400: 'Denderwindeke',
                                                   9401: 'Pollare', 9402: 'Meerbeke', 9403: 'Neigem', 9404: 'Aspelare',
                                                   9406: 'Outer', 9420: 'Burst', 9450: 'Heldergem', 9451: 'Kerksken',
                                                   9470: 'Denderleeuw', 9472: 'Iddergem', 9473: 'Welle',
                                                   9500: 'Zarlardinge', 9506: 'Schendelbeke',
                                                   9520: 'Sint-Lievens-Houtem', 9521: 'Letterhoutem',
                                                   9550: 'Sint-lievens-esse', 9551: 'Ressegem', 9552: 'Borsbeke',
                                                   9570: 'Lierde', 9571: 'Hemelveerdegem', 9572: 'Sint-martens-lierde',
                                                   9600: 'Ronse', 9620: 'Sint-Maria-Oudenhove', 9630: 'Munkzwalm',
                                                   9636: 'Nederzwalm-Hermelgem', 9660: 'Everbeek', 9661: 'Parike',
                                                   9667: 'Horebeke', 9680: 'Maarkedal', 9681: 'Nukerke',
                                                   9688: 'Schorisse', 9690: 'Berchem', 9700: 'Oudenaarde',
                                                   9750: 'Ouwegem', 9770: 'Kruishoutem', 9771: 'Nokere',
                                                   9772: 'Wannegem-lede', 9790: 'Moregem', 9800: 'Meigem', 9810: 'Eke',
                                                   9820: 'Schelderode', 9830: 'Sint-martens-latem', 9831: 'Deurle',
                                                   9840: 'Zevergem', 9850: 'Hansbeke', 9860: 'Scheldewindeke',
                                                   9870: 'Olsene', 9880: 'Lotenhulle', 9881: 'Bellem', 9890: 'Vurste',
                                                   9900: 'Eeklo', 9910: 'Ursel', 9920: 'Lovendegem',
                                                   9921: 'Vinderhoute', 9930: 'Zomergem', 9931: 'Oostwinkel',
                                                   9932: 'Ronsele', 9940: 'Ertvelde', 9950: 'Waarschoot',
                                                   9960: 'Assenede', 9961: 'Boekhoute', 9968: 'Bassevelde',
                                                   9970: 'Kaprijke', 9971: 'Lembeke', 9980: 'Sint-Laureins',
                                                   9981: 'Sint-Margriete', 9982: 'Sint-Jan-In-Eremo',
                                                   9988: 'Watervliet', 9990: 'Maldegem', 9991: 'Adegem',
                                                   9992: 'Middelburg', 1500: 'Halle', 1501: 'Buizingen',
                                                   1502: 'Lembeek', 1540: 'Herne', 1541: 'Sint-pieters-kapelle',
                                                   1547: 'Bever', 1560: 'Hoeilaart', 1570: 'Tollembeek',
                                                   1600: 'Sint-Pieters-Leeuw', 1601: 'Ruisbroek', 1602: 'Vlezenbeek',
                                                   1620: 'Drogenbos', 1630: 'Linkebeek', 1640: 'Sint-Genesius-Rode',
                                                   1650: 'Beersel', 1651: 'Lot', 1652: 'Alsemberg', 1653: 'Dworp',
                                                   1654: 'Huizingen', 1670: 'Heikruis', 1671: 'Elingen', 1673: 'Beert',
                                                   1674: 'Bellingen', 1700: 'Sint-ulriks-kapelle', 1701: 'Itterbeek',
                                                   1702: 'Groot-bijgaarden', 1703: 'Schepdaal', 1730: 'Asse',
                                                   1731: 'Relegem', 1740: 'Ternat', 1741: 'Wambeek',
                                                   1742: 'Sint-Katherina-Lombeek', 1745: 'Opwijk',
                                                   1750: 'Sint-martens-lennik', 1755: 'Kester', 1760: 'Pamel',
                                                   1761: 'Borchtlombeek', 1770: 'Liedekerke', 1780: 'Wemmel',
                                                   1785: 'Brussegem', 1790: 'Essene', 1800: 'Peutie',
                                                   1820: 'Steenokkerzeel', 1830: 'Machelen', 1831: 'Diegem',
                                                   1840: 'Londerzeel', 1850: 'Grimbergen', 1851: 'Humbeek',
                                                   1852: 'Beigem', 1853: 'Strombeek-bever', 1860: 'Meise',
                                                   1861: 'Wolvertem', 1880: 'Kapelle-op-Den-Bos', 1910: 'Berg',
                                                   1930: 'Nossegem', 1932: 'Sint-Stevens-Woluwe', 1933: 'Sterrebeek',
                                                   1950: 'Kraainem', 1970: 'Wezembeek-Oppem', 1980: 'Zemst',
                                                   1981: 'Hofstade', 1982: 'Elewijt', 3000: 'Leuven', 3001: 'Heverlee',
                                                   3010: 'Kessel-Lo', 3012: 'Wilsele', 3018: 'Wijgmaal',
                                                   3020: 'Winksele', 3040: 'Sint-agatha-rode', 3050: 'Oud-Heverlee',
                                                   3051: 'Sint-Joris-Weert', 3052: 'Blanden', 3053: 'Haasrode',
                                                   3054: 'Vaalbeek', 3060: 'Korbeek-Dijle', 3061: 'Leefdaal',
                                                   3070: 'Kortenberg', 3071: 'Erps-kwerps', 3078: 'Everberg',
                                                   3080: 'Tervuren', 3090: 'Overijse', 3110: 'Rotselaar',
                                                   3111: 'Wezemaal', 3118: 'Werchter', 3120: 'Tremelo', 3128: 'Baal',
                                                   3130: 'Betekom', 3140: 'Keerbergen', 3150: 'Haacht',
                                                   3190: 'Boortmeerbeek', 3191: 'Hever', 3200: 'Aarschot',
                                                   3201: 'Langdorp', 3202: 'Rillaar', 3210: 'Linden', 3211: 'Binkom',
                                                   3212: 'Pellenberg', 3220: 'Sint-pieters-rode', 3221: 'Nieuwrode',
                                                   3270: 'Scherpenheuvel-Zichem', 3271: 'Zichem', 3272: 'Testelt',
                                                   3290: 'Webbekom', 3293: 'Kaggevinne', 3294: 'Molenstede',
                                                   3300: 'Bost', 3320: 'Meldert', 3321: 'Outgaarden',
                                                   3350: 'Neerhespen', 3360: 'Korbeek-Lo', 3370: 'Vertrijk',
                                                   3380: 'Glabbeek-Zuurbemde', 3381: 'Kapellen', 3384: 'Attenrode',
                                                   3390: 'Sint-Joris-Winge', 3391: 'Meensel-Kiezegem', 3400: 'Landen',
                                                   3401: 'Wezeren', 3404: 'Attenhoven', 3440: 'Zoutleeuw',
                                                   3450: 'Grazen', 3454: 'Rummen', 3460: 'Assent',
                                                   3461: 'Molenbeek-Wersbeek', 3470: 'Sint-margriete-houtem',
                                                   3471: 'Hoeleden', 3472: 'Kersbeek-miskom', 3473: 'Waanrode',
                                                   1300: 'Limal', 1301: 'Bierges', 1310: 'La hulpe', 1315: 'Opprebais',
                                                   1320: "L'ecluse", 1325: 'Bonlez', 1330: 'Rixensart',
                                                   1331: 'Rosières', 1332: 'Genval', 1340: 'Ottignies-louvain-la-neuve',
                                                   1341: 'Céroux-mousty', 1342: 'Limelette', 1348: 'Louvain-la-neuve',
                                                   1350: 'Orp-jauche', 1357: 'Neerheylissem',
                                                   1360: 'Thorembais-les-béguines', 1367: 'Bomal',
                                                   1370: 'Jodoigne-souveraine', 1380: 'Plancenoit', 1390: 'Nethen',
                                                   1400: 'Monstreux', 1401: 'Baulers', 1402: 'Thines', 1404: 'Bornival',
                                                   1410: 'Waterloo', 1420: "Braine-l'alleud",
                                                   1421: 'Ophain-bois-seigneur-isaac', 1428: 'Lillois-witterzée',
                                                   1430: 'Rebecq', 1435: 'Corbais', 1440: 'Braine-le-château',
                                                   1450: 'Chastre-villeroux-blanmont', 1457: 'Walhain-saint-paul',
                                                   1460: 'Virginal-samme', 1461: 'Haut-ittre', 1470: 'Genappe',
                                                   1471: 'Loupoigne', 1472: 'Vieux-genappe', 1473: 'Glabais',
                                                   1474: 'Ways', 1476: 'Houtain-le-val', 1480: 'Clabecq',
                                                   1490: 'Court-saint-etienne', 1495: 'Sart-dames-avelines',
                                                   8000: 'Koolkerke', 8020: 'Hertsberge', 8200: 'Sint-andries',
                                                   8210: 'Veldegem', 8211: 'Aartrijke', 8300: 'Knokke-Heist',
                                                   8301: 'Heist-aan-zee', 8310: 'Assebroek', 8340: 'Lapscheure',
                                                   8370: 'Uitkerke', 8377: 'Zuienkerke', 8380: 'Lissewege',
                                                   8400: 'Stene', 8420: 'Wenduine', 8421: 'Vlissegem',
                                                   8430: 'Middelkerke', 8431: 'Wilskerke', 8432: 'Leffinge',
                                                   8433: 'Slijpe', 8434: 'Westende', 8450: 'Bredene', 8460: 'Roksem',
                                                   8470: 'Snaaskerke', 8480: 'Ichtegem', 8490: 'Zerkegem',
                                                   8500: 'Kortrijk', 8501: 'Heule', 8510: 'Rollegem', 8511: 'Aalbeke',
                                                   8520: 'Kuurne', 8530: 'Harelbeke', 8531: 'Hulste', 8540: 'Deerlijk',
                                                   8550: 'Zwevegem', 8551: 'Heestert', 8552: 'Moen', 8553: 'Otegem',
                                                   8554: 'Sint-Denijs', 8560: 'Wevelgem', 8570: 'Anzegem',
                                                   8572: 'Kaster', 8573: 'Tiegem', 8580: 'Avelgem', 8581: 'Kerkhove',
                                                   8582: 'Outrijve', 8583: 'Bossuit', 8587: 'Helkijn', 8600: 'Keiem',
                                                   8610: 'Werken', 8620: 'Nieuwpoort', 8630: 'Wulveringem',
                                                   8640: 'Oostvleteren', 8647: 'Pollinkhove', 8650: 'Houthulst',
                                                   8660: 'De panne', 8670: 'Koksijde', 8680: 'Koekelare',
                                                   8690: 'Sint-Rijkers', 8691: 'Izenberge', 8700: 'Tielt',
                                                   8710: 'Wielsbeke', 8720: 'Oeselgem', 8730: 'Beernem', 8740: 'Pittem',
                                                   8750: 'Zwevezele', 8755: 'Ruiselede', 8760: 'Meulebeke',
                                                   8770: 'Ingelmunster', 8780: 'Oostrozebeke', 8790: 'Waregem',
                                                   8791: 'Beveren', 8792: 'Desselgem', 8793: 'Sint-Eloois-Vijve',
                                                   8800: 'Oekene', 8810: 'Lichtervelde', 8820: 'Torhout',
                                                   8830: 'Hooglede', 8840: 'Oostnieuwkerke', 8850: 'Ardooie',
                                                   8851: 'Koolskamp', 8860: 'Lendelede', 8870: 'Kachtem',
                                                   8880: 'Sint-eloois-winkel', 8890: 'Dadizele', 8900: 'Sint-jan',
                                                   8902: 'Voormezele', 8904: 'Zuidschote', 8906: 'Elverdinge',
                                                   8908: 'Vlamertinge', 8920: 'Bikschote', 8930: 'Lauwe',
                                                   8940: 'Wervik', 8950: 'Nieuwkerke', 8951: 'Dranouter',
                                                   8952: 'Wulvergem', 8953: 'Wijtschate', 8954: 'Westouter',
                                                   8956: 'Kemmel', 8957: 'Mesen', 8958: 'Loker', 8970: 'Poperinge',
                                                   8972: 'Roesbrugge-haringe', 8978: 'Watou', 8980: 'Geluveld'}


def belgian_postal_code(list_with_postal_codes):
    '''
    Return the postal code and city name given the postal code. It is also possible that there are already city names
    in the given list, then it adds this city name to the resulting list without editing it.
    '''
    list_result = []

    for code in list_with_postal_codes:
        if code is None:
            list_result.append(None)
        elif not isinstance(code, int):
            list_result.append(code)
        else:
            list_result.append(str(code) + ' ' + dictionary_belgian_postal_codes_with_city_names[code])

    return list_result


def webscrape(path_workbook, name_worksheet, column1, column2, column1_insert, column2_insert, beginning_row,
              ending_row):
    '''

    :param path_workbook: path to the excel file that needs to be edited
    :param name_worksheet: name inside the excel workbook of the worksheet that needs to be edited
    :param column1: column where the first postal codes are stored
    :param column2: column where the second postal codes are stored
    :param column: column where the result of the program needs to be written to
    '''

    # Takes the two lists with the needed postal codes or city names(from an to)
    tuple_with_values = excel.get_values_columns(path_workbook, name_worksheet, column1, column2, beginning_row,
                                                 ending_row)
    list_with_postal_codes1 = tuple_with_values[0]
    list_with_postal_codes2 = tuple_with_values[1]

    list_with_postal_codes_and_city_names1 = belgian_postal_code(list_with_postal_codes1)
    list_with_postal_codes_and_city_names2 = belgian_postal_code(list_with_postal_codes2)

    # Open Google maps website
    driver = webdriver.Chrome(
        executable_path=r'C:\Users\jensb\OneDrive\Documenten\Programming\Python\renewi_distance_calculator_excel\templates\chromedriver.exe')
    driver.get('https://www.google.be/maps/dir///@50.8782126,4.6874076,14z/data=!4m2!4m1!3e0')

    time.sleep(1)

    # Find the xpath to the search bars to enter the postal codes
    xpath_first_search_bar = '/html/body/jsl/div[3]/div[9]/div[3]/div[1]/div[2]/div/div[3]/div[1]/div[1]/div[2]/div/div/input'
    xpath_second_search_bar = '/html/body/jsl/div[3]/div[9]/div[3]/div[1]/div[2]/div/div[3]/div[1]/div[2]/div[2]/div/div/input'

    list_distances = []
    list_times = []


    for index in range(len(list_with_postal_codes_and_city_names1)):


        not_found = False

        # Delete text in search boxes and enter the postal codes
        driver.find_element_by_xpath(xpath_first_search_bar).send_keys(Keys.CONTROL + 'a')
        driver.find_element_by_xpath(xpath_first_search_bar).send_keys(Keys.DELETE)

        driver.find_element_by_xpath(xpath_second_search_bar).send_keys(Keys.CONTROL + 'a')
        driver.find_element_by_xpath(xpath_second_search_bar).send_keys(Keys.DELETE)

        city1 = list_with_postal_codes_and_city_names1[index]
        city2 = list_with_postal_codes_and_city_names2[index]

        if city1 is None or city2 is None:
            list_distances.append('null')
            list_times.append('null')
        elif city1 == city2:
            list_distances.append(0)
            list_times.append(0)
        else:
            # Enter the postal codes into the search bars
            driver.find_element_by_xpath(xpath_first_search_bar).send_keys(city1)
            driver.find_element_by_xpath(xpath_second_search_bar).send_keys(city2)

            driver.find_element_by_xpath(xpath_second_search_bar).send_keys(Keys.ENTER)

            # Get page html
            page_soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Check if the spinner symbol is still loading

            while page_soup.find('div', {'class': 'section-loading-spinner'}):
                page_soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Wait until the box with the traject information has appeared
            while not page_soup.find('div', {'class': 'section-directions-trip-numbers'}) and not not_found:
                # If google maps can't calculate a route, add 'null' to the lists.
                length_widget_error = len(page_soup.findAll('h2', {'class': 'widget-directions-error'}))
                length_section_error = len(page_soup.findAll('div', {'class': 'section-directions-error-primary-text'}))
                if length_widget_error != 0 or length_section_error != 0:
                    not_found = True

                page_soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Get all the boxes of the possible trajects from one city to another
            containers = page_soup.findAll('div', {'class': 'section-directions-trip-numbers'})

            if not not_found:
                if len(containers) == 0:
                    list_distances.append('null')
                    list_times.append('null')
                else:
                    container = containers[0]

                    time_box = container.findAll('div', {'class': 'section-directions-trip-duration'})[0]
                    distance_box = container.findAll('div', {
                        'class': 'section-directions-trip-distance section-directions-trip-secondary-text'})[0]

                    travel_time = time_box.findAll('span')[0].text
                    distance = distance_box.findAll('div')[0].text

                    list_times.append(travel_time)
                    list_distances.append(distance)
            else:
                list_times.append('not_found')
                list_distances.append('not_found')

    excel.insert_values_column(path_workbook, name_worksheet, column1_insert, list_distances, beginning_row)
    excel.insert_values_column(path_workbook, name_worksheet, column2_insert, list_times, beginning_row)

    driver.close()
