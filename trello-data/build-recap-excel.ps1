# Build C:\Users\claud\Downloads\livret-accueil\recap-apparts.xlsx
# 22 apparts × ~40 colonnes — données complètes pour vérif Claudine

$ErrorActionPreference = 'Stop'

# ==== Données (22 apparts, ordre logique par résidence) ====
$apparts = @(
  # === Le 23 Maréchal Juin ===
  @{ slug='kitesurf'; nom='Kitesurf'; com='Kitesurf'; adr='23 rue Maréchal Juin, 62600 Berck'; res='Le 23'; zone='marechal';
     type='Appartement'; surface='34 m²'; lits='4 couchages · chambre 160×200 · canapé gigogne 2×(90×190)'; etage='2e étage'; ascenseur='Non'; parking='Dans la rue';
     videoAcces='https://drive.google.com/file/d/1yoVLVnDSSuuU90lgiMc3-EHvn3vVVqLv/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Non'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='2';
     wifiSsid='BBOX - 4929C743'; wifiPwd='KXv9 DNH2 uDd5 n165 1u';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Oui';
     prestataire='Rémi'; entite='LMP'; photos=2 },

  @{ slug='hamac'; nom='Hamac'; com='Hamac'; adr='23 rue Maréchal Juin, 62600 Berck'; res='Le 23'; zone='marechal';
     type='Appartement'; surface='35 m²'; lits='4 couchages · chambre 160×200 · canapé convertible 140×190'; etage='2e étage'; ascenseur='Non'; parking='Dans la rue';
     videoAcces='https://drive.google.com/file/d/1FsKU-XGly3eeHe0fYpDQ3sFK3GM3jpBg/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Oui'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='2';
     wifiSsid='BBOX-4DA61D'; wifiPwd='9LMRXwLjC3qRxMLU93';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Non (lit double + canapé-lit)';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Oui';
     prestataire='Rémi'; entite='LMP'; photos=0 },

  @{ slug='paddle'; nom='Paddle'; com='Paddle'; adr='23 rue Maréchal Juin, 62600 Berck'; res='Le 23'; zone='marechal';
     type='Appartement'; surface='30 m²'; lits='4 couchages · chambre 160×200 · canapé gigogne 2×(90×190)'; etage='2e étage'; ascenseur='Non'; parking='Dans la rue';
     videoAcces='https://drive.google.com/file/d/1lRwnNiII1xYEauzS5ddg6J4jKLQm9uum/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Oui'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='2';
     wifiSsid='BBOX-39A25D'; wifiPwd='Twuy7AjqMpFAY3bSTM';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Oui';
     prestataire='Rémi'; entite='LMP'; photos=0 },

  @{ slug='surf'; nom='Surf'; com='Surf'; adr='23 rue Maréchal Juin, 62600 Berck'; res='Le 23'; zone='marechal';
     type='Appartement'; surface='30 m²'; lits='4 couchages · chambre 160×200 · canapé gigogne 2×(90×190)'; etage='1er étage'; ascenseur='Non'; parking='Dans la rue';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Oui'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='2';
     wifiSsid='BBOX-1A4DA8'; wifiPwd='To check';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Rémi'; entite='LMP'; photos=0 },

  @{ slug='famille'; nom='Famille'; com='3 Chambres 100m Plage'; adr='23 rue Maréchal Juin, 62600 Berck'; res='Le 23'; zone='marechal';
     type='Grand appartement 3 chambres'; surface='70 m²'; lits='6+ couchages · 3 chambres'; etage='À préciser'; ascenseur='Non'; parking='Dans la rue';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Induction'; frigo='Frigo + congélateur séparé'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Non'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='2';
     wifiSsid='Famille-Berck'; wifiPwd='To check';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Rémi'; entite='LMP'; photos=6 },

  # === Le 25 Maréchal Juin ===
  @{ slug='balneo'; nom='Balnéo'; com='Balnéo Garden'; adr='25 rue Maréchal Juin, 62600 Berck'; res='Le 25'; zone='marechal';
     type='Appartement'; surface='53 m²'; lits='6 couchages · chambre 1: 160×200 · chambre 2: 2×(90×190) · canapé 140×190'; etage='RDC'; ascenseur='Non'; parking='✅ Parking sécurisé inclus';
     videoAcces='https://drive.google.com/file/d/1tc8QjWR0X3tYzfez82oHtJ4DMdQ2uLhw/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Frigo américain (congélateur intégré)'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Oui'; douche='Oui';
     balneo='Cetana (2 places face à face)'; sauna='Non';
     balcon='Oui'; brasero='Oui'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='BBOX-CE0F6E'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Clara'; entite='LMP'; photos=6 },

  @{ slug='cocon'; nom='Cocon Romantique'; com='Cocon Romantique'; adr='25 rue Maréchal Juin, 62600 Berck'; res='Le 25'; zone='marechal';
     type='Appartement'; surface='40 m²'; lits='2 couchages · chambre 160×200'; etage='RDC'; ascenseur='Non'; parking='Dans la rue';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Oui'; douche='Oui';
     balneo='Cetana (2 places face à face)'; sauna='Oui';
     balcon='Non'; brasero='Oui'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Cocon-Berck'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Clara'; entite='LMP'; photos=6 },

  # === Terrasse / rue des Sables ===
  @{ slug='terrasse'; nom='Terrasse'; com='La Terrasse'; adr='17 rue du Grand Hôtel, 62600 Berck'; res='Terrasse'; zone='terrasse';
     type='Grand appartement'; surface='63 m²'; lits='8 couchages · chambre 1: 160×200 · chambre 2: 140×200 + 2×(90×190) · canapé 140×190'; etage='RDC'; ascenseur='Non'; parking='Dans la rue';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Frigo + congélateur séparé'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Oui'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Oui'; brasero='Non'; bbq='Oui';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Terrasse-Berck'; wifiPwd='To check';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='✅ Lit bébé + chaise haute'; palier23='Non';
     prestataire='Christelle'; entite='LMP'; photos=0 },

  @{ slug='maisonnette'; nom='Maisonnette'; com='Cœur des Sables'; adr='2 rue des Sables, 62600 Berck'; res='Terrasse'; zone='terrasse';
     type='Studio'; surface='23 m²'; lits='2 lits · 160×200'; etage='RDC'; ascenseur='Non'; parking='Dans la rue';
     videoAcces='https://drive.google.com/file/d/14nBlb8KgwoHIzeLN7z8mi9n9eAP6CuAq/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo';
     ll='Oui'; lv='Non'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Oui'; brasero='Oui'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Maisonnette-Berck'; wifiPwd='To check';
     tarifOption='15 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Christelle'; entite='LMP'; photos=6 },

  # === Face Mer ===
  @{ slug='facemer'; nom='Face Mer'; com='Face Mer Esplanade'; adr='1 rue de la Plage, 62600 Berck'; res='Face Mer'; zone='esplanade';
     type='Appartement'; surface='À préciser'; lits='À préciser'; etage='2e étage'; ascenseur='Non'; parking='Dans la rue';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Frigo + congélateur séparé'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Oui (vue mer)'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='FaceMer-Berck'; wifiPwd='To check';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='✅ Lit bébé + chaise haute'; palier23='Non';
     prestataire='Christelle'; entite='LMP'; photos=6 },

  # === Grand Large ===
  @{ slug='grandlarge'; nom='Grand Large'; com='Hamac Face Mer'; adr='45 Esplanade Parmentier, 62600 Berck'; res='Grand Large'; zone='esplanade';
     type='Appartement'; surface='38 m²'; lits='4 couchages · chambre 160×200 · canapé convertible 140×190'; etage='3e étage'; ascenseur='Non'; parking='Dans la rue';
     videoAcces='https://drive.google.com/file/d/1yT8uTNvdaBE5eM7y-VJIBBW5zskDZLfD/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Oui (hamac)'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='GrandLarge-Berck'; wifiPwd='To check';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Christelle'; entite='SAS Maréchal Invest'; photos=0 },

  # === Apollo (2 av Marianne) ===
  @{ slug='apolove'; nom='Apolove'; com='Suite Romantique Face Mer'; adr='2 Avenue Marianne (studio), 62600 Berck'; res='Apollo'; zone='esplanade';
     type='Studio'; surface='25 m²'; lits='2 couchages · chambre 160×200'; etage='5e étage'; ascenseur='Oui'; parking='✅ Parking couvert n°206 inclus (résidence Albatros, rue Saint-Louis)';
     videoAcces='https://drive.google.com/file/d/1AFgKVhahrxvljkL7vVa8ZU1-zA0BgHwo/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Non'; lv='Oui'; baignoire='Oui'; douche='Oui';
     balneo='Indra Wave (compact, 2 personnes face à face)'; sauna='Non';
     balcon='Oui'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='suite2'; wifiPwd='2Avmarianne';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='❌ Interdits'; bebe='Non'; palier23='Non';
     prestataire='Clara'; entite='SAS Maréchal Invest'; photos=6 },

  @{ slug='apollo'; nom='Apollo'; com='Panoramique 2 Chambres'; adr='2 Avenue Marianne (T3), 62600 Berck'; res='Apollo'; zone='esplanade';
     type='Appartement'; surface='53 m²'; lits='6 couchages · 2 chambres 160×200 · canapé 140×190'; etage='1er étage'; ascenseur='Oui'; parking='✅ Parking n°502 inclus (portail avec bip)';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Frigo américain (congélateur intégré)'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Oui (vue panoramique)'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Apollo-Berck'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='❌ Interdits'; bebe='Non'; palier23='Non';
     prestataire='Clara'; entite='SAS Maréchal Invest'; photos=0 },

  @{ slug='kingston'; nom='Kingston'; com='Face Mer Exceptionnel'; adr='2 Avenue Marianne (appt 16, RDC), 62600 Berck'; res='Apollo'; zone='esplanade';
     type='Appartement'; surface='47 m²'; lits='4 couchages · canapé lit 160'; etage='RDC'; ascenseur='Non'; parking='✅ Parking n°16 inclus (juste devant)';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Frigo + congélateur séparé'; cafetiere='Senseo';
     ll='Oui'; lv='Non'; baignoire='Oui'; douche='Non';
     balneo='Non'; sauna='Non';
     balcon='Oui'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Kingston-Berck'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Clara'; entite='SCI'; photos=1 },

  @{ slug='reserve'; nom='La Réserve'; com='La Réserve'; adr='2 Avenue Marianne, 62600 Berck'; res='Apollo'; zone='esplanade';
     type='Appartement'; surface='À préciser'; lits='À préciser'; etage='À préciser'; ascenseur='Oui'; parking='À préciser';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='À préciser'; lv='À préciser'; baignoire='À préciser'; douche='À préciser';
     balneo='Non'; sauna='Non';
     balcon='À préciser'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='À préciser'; wifiPwd='À préciser';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Clara'; entite='SAS Maréchal Invest'; photos=6 },

  # === Albatros ===
  @{ slug='albatros'; nom='Albatros'; com='Hamac Chic Vue Mer'; adr='10 Avenue Marianne, 62600 Berck'; res='Albatros'; zone='esplanade';
     type='Appartement'; surface='25 m²'; lits='4 couchages · chambre 140×200 · canapé convertible 140×190'; etage='3e étage'; ascenseur='Oui'; parking='Dans la rue';
     videoAcces='https://drive.google.com/file/d/1IBQUDOQNG4ENcgeAtUPYJInJFTUpWpgM/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Non (local à vélos)'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Albatros-Berck'; wifiPwd='To check';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Clara'; entite='SAS Maréchal Invest'; photos=0 },

  # === Esplanade (59 Esplanade Parmentier) ===
  @{ slug='miniloveroom'; nom='Mini Love Room'; com='Balnéo Face Mer'; adr='59 Esplanade Parmentier, 62600 Berck'; res='Esplanade'; zone='esplanade';
     type='Studio'; surface='21 m²'; lits='2 couchages · lit 160×200'; etage='3e étage'; ascenseur='Non'; parking='Dans la rue';
     videoAcces='https://drive.google.com/file/d/1AWX2lXeXVBwGkngMM33dsRWiLXG1nny_/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Aucun (juste micro-ondes grill)'; cafetiere='Senseo';
     ll='Non'; lv='Non'; baignoire='Oui'; douche='Oui';
     balneo='Indra Wave (compact, 2 personnes face à face)'; sauna='Non';
     balcon='Non'; brasero='Non'; bbq='Non';
     tv='Vidéoprojecteur + Molotov Premium'; remotes='1';
     wifiSsid='MiniLove-Berck'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Christelle'; entite='SAS Maréchal Invest'; photos=6 },

  @{ slug='grandeloveroom'; nom='Grande Love Room'; com='Romantique Face Mer'; adr='59 Esplanade Parmentier, 62600 Berck'; res='Esplanade'; zone='esplanade';
     type='Studio'; surface='28 m²'; lits='2 couchages · lit 160×200'; etage='3e étage'; ascenseur='Non'; parking='Dans la rue';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo';
     ll='Non'; lv='Oui'; baignoire='Non'; douche='Oui';
     balneo='Indra Wave (compact, 2 personnes face à face)'; sauna='Non';
     balcon='Non'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='GrandeLove-Berck'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='❌ Interdits'; bebe='Non'; palier23='Non';
     prestataire='Christelle'; entite='SAS Maréchal Invest'; photos=6 },

  # === Rue Jeanne ===
  @{ slug='jeanne'; nom='Jeanne'; com='Cœur de Berck'; adr='23 Rue Jeanne, 62600 Berck'; res='Rue Jeanne'; zone='jeanne';
     type='Appartement'; surface='48 m²'; lits='4 couchages · chambre 160×200 + canapé 140×190'; etage='1er étage · appartement 3'; ascenseur='Non'; parking='Dans la rue + parking stade gratuit';
     videoAcces='https://drive.google.com/file/d/1cQR6Hv6br1B7-VEVXDFCPFT_Z3H7OL75/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Gaz'; frigo='Frigo + congélateur séparé'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Non'; baignoire='Non'; douche='Oui';
     balneo='Non'; sauna='Non';
     balcon='Non'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='TP-LINK _6322'; wifiPwd='83968280';
     tarifOption='15 €'; lingeInclus='Non (en option)'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Christelle'; entite='SAS Maréchal Invest'; photos=6 },

  @{ slug='evasion'; nom='Évasion'; com="L'Évasion Balnéo & Sauna"; adr='Rue Jeanne, 62600 Berck'; res='Rue Jeanne'; zone='jeanne';
     type='Appartement'; surface='38 m²'; lits='4 couchages · chambre 160×200 + canapé 140×190'; etage='1er étage'; ascenseur='Non'; parking='Dans la rue + parking stade gratuit';
     videoAcces='https://drive.google.com/file/d/1UGxYQjz06UbgdcJZw5Zn0lZ0a_AmTa0j/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Oui'; douche='Oui';
     balneo='Cetana (2 places face à face)'; sauna='Oui';
     balcon='Non'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Evasion-Berck'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Christelle'; entite='SAS Maréchal Invest'; photos=0 },

  @{ slug='rotonde'; nom='Rotonde'; com='La Rotonde du Bien-être'; adr='Rue Jeanne, 62600 Berck'; res='Rue Jeanne'; zone='jeanne';
     type='Appartement'; surface='55 m²'; lits='5 couchages · chambre 1: 160×200 · chambre 2: 2×(90×190) · canapé 140×190'; etage='RDC'; ascenseur='Non'; parking='Dans la rue + parking stade gratuit';
     videoAcces=''; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Oui'; douche='Oui';
     balneo='Cetana (2 places face à face)'; sauna='Oui (max 65°)';
     balcon='Non'; brasero='Non'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Rotonde-Berck'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Christelle'; entite='SAS Maréchal Invest'; photos=0 },

  @{ slug='patio'; nom='Patio'; com='Patio Romantique Chic'; adr='Rue Jeanne, 62600 Berck'; res='Rue Jeanne'; zone='jeanne';
     type='Appartement'; surface='47 m²'; lits='6 couchages · chambre 1: 160×200 · chambre 2: 2×(90×190) · canapé 140×190'; etage='RDC'; ascenseur='Non'; parking='Dans la rue + parking stade gratuit';
     videoAcces='https://drive.google.com/file/d/1DcKUvgC95o72FmCf8Wc-KeCE-chpgV4I/view'; videoParking='https://www.youtube.com/watch?v=HVP1-x_L478';
     plaque='Vitro-céramique'; frigo='Compartiment freezer'; cafetiere='Senseo + filtre';
     ll='Oui'; lv='Oui'; baignoire='Oui'; douche='Oui';
     balneo='Cetana (2 places face à face)'; sauna='Non';
     balcon='Oui (patio)'; brasero='Oui'; bbq='Non';
     tv='Molotov standard'; remotes='1';
     wifiSsid='Patio-Berck'; wifiPwd='To check';
     tarifOption='20 €'; lingeInclus='Oui'; packLitSimple='Oui';
     chiens='Oui (sur demande)'; bebe='Non'; palier23='Non';
     prestataire='Christelle'; entite='SAS Maréchal Invest'; photos=6 }
)

# Colonnes (en ordre logique)
$cols = @(
  '#','Slug','Nom court','Nom commercial (site)','Adresse','Résidence','Zone',
  'Type logement','Surface','Couchages','Étage','Ascenseur','Parking',
  'Vidéo accès','Vidéo parking',
  'Plaque cuisson','Frigo / Congélation','Cafetière',
  'Lave-linge','Lave-vaisselle','Baignoire','Douche',
  'Balnéo Spalina','Sauna',
  'Balcon / Terrasse','Brasero','Barbecue',
  'TV','Nb télécommandes',
  'WiFi réseau','WiFi mot de passe',
  'Tarif check-in/out','Linge inclus','Pack lit simple',
  'Chiens','Lit bébé / chaise haute','Palier Le 23',
  'Prestataire ménage','Entité juridique','Photos disponibles'
)

# Création Excel
Write-Host "Lancement Excel COM..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$ws = $wb.ActiveSheet
$ws.Name = 'Récap 22 apparts'

# Header
for ($c = 0; $c -lt $cols.Count; $c++) {
  $ws.Cells.Item(1, $c+1) = $cols[$c]
}

# Style header
$range = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1, $cols.Count))
$range.Font.Bold = $true
$range.Font.Size = 11
$range.Interior.Color = 0x101510  # noir/marron
$range.Font.Color = 0x5EAFD4  # or
$range.HorizontalAlignment = -4108  # center
$range.RowHeight = 32

# Lignes
for ($i = 0; $i -lt $apparts.Count; $i++) {
  $a = $apparts[$i]
  $row = $i + 2
  $vals = @(
    ($i+1), $a.slug, $a.nom, $a.com, $a.adr, $a.res, $a.zone,
    $a.type, $a.surface, $a.lits, $a.etage, $a.ascenseur, $a.parking,
    $a.videoAcces, $a.videoParking,
    $a.plaque, $a.frigo, $a.cafetiere,
    $a.ll, $a.lv, $a.baignoire, $a.douche,
    $a.balneo, $a.sauna,
    $a.balcon, $a.brasero, $a.bbq,
    $a.tv, $a.remotes,
    $a.wifiSsid, $a.wifiPwd,
    $a.tarifOption, $a.lingeInclus, $a.packLitSimple,
    $a.chiens, $a.bebe, $a.palier23,
    $a.prestataire, $a.entite, $a.photos
  )
  for ($c = 0; $c -lt $vals.Count; $c++) {
    $ws.Cells.Item($row, $c+1) = $vals[$c]
  }
}

# Largeurs colonnes auto
$ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($apparts.Count+1, $cols.Count)).EntireColumn.AutoFit() | Out-Null

# Wrap text + freeze top row
$ws.Range($ws.Cells.Item(2,1), $ws.Cells.Item($apparts.Count+1, $cols.Count)).WrapText = $true
$ws.Range($ws.Cells.Item(2,1), $ws.Cells.Item($apparts.Count+1, $cols.Count)).VerticalAlignment = -4160 # top
$ws.Application.ActiveWindow.SplitRow = 1
$ws.Application.ActiveWindow.FreezePanes = $true

# Plafond largeur (pour pas exploser)
foreach ($col in 1..$cols.Count) {
  $w = $ws.Columns.Item($col).ColumnWidth
  if ($w -gt 35) { $ws.Columns.Item($col).ColumnWidth = 35 }
  if ($w -lt 12) { $ws.Columns.Item($col).ColumnWidth = 12 }
}

# Filtres auto
$ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($apparts.Count+1, $cols.Count)).AutoFilter() | Out-Null

# Save
$out = 'C:\Users\claud\Downloads\livret-accueil\recap-apparts.xlsx'
if (Test-Path $out) { Remove-Item $out -Force }
$wb.SaveAs($out, 51) # xlOpenXMLWorkbook = 51

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "OK : $out cree avec $($apparts.Count) apparts et $($cols.Count) colonnes"
