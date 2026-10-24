-- Mode démo uniquement : équipe fictive, quelques cartes et soldes de départ.
insert into auth.users(id, email, raw_user_meta_data) values
  ('00000000-0000-0000-0000-000000000001', 'gerante@ayael.app', '{"full_name":"Gérante (démo)"}'),
  ('00000000-0000-0000-0000-000000000002', 'agent1@ayael.app',  '{"full_name":"Agent 1 (démo)"}'),
  ('00000000-0000-0000-0000-000000000003', 'agent2@ayael.app',  '{"full_name":"Agent 2 (démo)"}')
on conflict do nothing;

update profiles set active = true where login in ('agent1','agent2');
update profiles set can_recharge = false where login = 'agent2';

select set_config('app.uid', '00000000-0000-0000-0000-000000000001', false);
select _caisse_of(id) from profiles;

select import_cards('{"cards":[
  {"card_number":"5027504","holder_name":"OLADELE RACHIDE TCHALA","bank":"GTCO"},
  {"card_number":"5051762","holder_name":"YASMINE DAOUDA","bank":"GTCO"},
  {"card_number":"5027658","holder_name":"CHARIFE ISSAKA","bank":"GTCO"},
  {"card_number":"5027437","holder_name":"TCHILABALO BALOUKIMONDOM M''BELOU","bank":"GTCO"},
  {"card_number":"23451460","holder_name":"KOUMANA MAYER BOUTORA","bank":"UBA"},
  {"card_number":"23451439","holder_name":"MASRA DOUMNAIBEI","bank":"UBA"},
  {"card_number":"18924902","holder_name":"DOPENYO AYAWA FLORENCE","bank":"ECOBANK","note":"4 derniers chiffres : 9316"}
]}');

update clients set phone = '90 11 22 33' where name = 'YASMINE DAOUDA';
update clients set phone = '91 22 33 44' where name = 'CHARIFE ISSAKA';
update settings set value = value || '{"slogan":"Le Spécialiste des Paiements en Ligne","address":"Sagbado, non loin de la Pharmacie Actuelle","phone":"+228 22 51 91 41","whatsapp":"+228 70 00 25 22","rccm":"TG-LOM 2020 B 2050","nif":"1001698268","website":"www.ayael-entreprise.com","footer":"Merci de votre confiance !"}' where key = 'company';

select op_ajustement(jsonb_build_object('account_id', id, 'target', v, 'details', 'Solde d''ouverture (démo)'))
from (values ('GTCO principal', 1810631), ('GTCO Ayael CI', 450000), ('UBA AfricSika', 296893), ('Ecobank', 120000),
             ('Mixx agent 1', 350000), ('Mixx agent 2', 280000), ('Flooz', 150000), ('Coffre', 500000)) s(n, v)
join accounts a on a.name = s.n;

select op_ajustement(jsonb_build_object('account_id', a.id, 'target', 50000, 'details', 'Fond de caisse (démo)'))
from accounts a where a.kind = 'caisse';

select set_config('app.uid', '00000000-0000-0000-0000-000000000002', false);
select op_caisse(jsonb_build_object('product_id', (select id from products where name = 'Dépôt Mixx'), 'received', 20000, 'change_given', 0, 'details', '93893071'));
select op_caisse(jsonb_build_object('product_id', (select id from products where name = 'Retrait Mixx'), 'amount', 3000, 'details', '92233123'));
select op_recharge(jsonb_build_object('gtp_id', (select id from accounts where name = 'GTCO principal'),
                                      'card_id', (select id from cards where card_number = '5051762'), 'amount', 100000));
select op_caisse(jsonb_build_object('product_id', (select id from products where name = 'Impression'), 'received', 1000, 'change_given', 500, 'details', 'Photocopies'));
select op_vente(jsonb_build_object(
  'client_id', (select id from clients where name = 'CHARIFE ISSAKA'), 'received', 50000,
  'lines', jsonb_build_array(
    jsonb_build_object('type', 'recharge', 'card_id', (select id from cards where card_number = '5027658'),
                       'gtp_id', (select id from accounts where name = 'GTCO principal'), 'amount', 42000),
    jsonb_build_object('type', 'caisse', 'product_id', (select id from products where name = 'Crédit de communication'), 'amount', 2000, 'details', 'Moov'))));

-- Deuxième agence : Noépé, avec son gérant, son agent et son propre compte Mixx.
select set_config('app.uid', '00000000-0000-0000-0000-000000000001', false);
select save_agence('{"name":"Noépé"}');
insert into auth.users(id, email, raw_user_meta_data) values
  ('00000000-0000-0000-0000-000000000004', 'gerant.noepe@ayael.app', '{"full_name":"Gérant Noépé (démo)"}'),
  ('00000000-0000-0000-0000-000000000005', 'agent.noepe@ayael.app',  '{"full_name":"Agent Noépé (démo)"}')
on conflict do nothing;
select save_profile(jsonb_build_object('id', '00000000-0000-0000-0000-000000000004', 'role', 'gerant', 'active', true,
                                       'agence_id', (select id from agences where name = 'Noépé')));
select save_profile(jsonb_build_object('id', '00000000-0000-0000-0000-000000000005', 'role', 'agent', 'active', true,
                                       'agence_id', (select id from agences where name = 'Noépé')));
select save_account(jsonb_build_object('name', 'Mixx Noépé', 'kind', 'electronique', 'agence_id', (select id from agences where name = 'Noépé')));
select op_ajustement(jsonb_build_object('account_id', (select id from accounts where name = 'Mixx Noépé'), 'target', 200000, 'details', 'Solde d''ouverture (démo)'));
select op_ajustement(jsonb_build_object('account_id', (select id from accounts where agent_id = '00000000-0000-0000-0000-000000000005'),
                                        'target', 30000, 'details', 'Fond de caisse (démo)'));
-- À Noépé : une recharge payée en espèces (Noépé doit 25 000) et un retrait RIA (la principale doit 10 000).
select set_config('app.uid', '00000000-0000-0000-0000-000000000005', false);
select op_vente(jsonb_build_object('client_id', (select id from clients where name = 'OLADELE RACHIDE TCHALA'), 'received', 25000,
  'lines', jsonb_build_array(jsonb_build_object('type', 'recharge', 'card_id', (select id from cards where card_number = '5027504'),
                                                'gtp_id', (select id from accounts where name = 'GTCO principal'), 'amount', 25000))));
select op_vente(jsonb_build_object('client_name', 'Client de passage',
  'lines', jsonb_build_array(jsonb_build_object('type', 'caisse', 'product_id', (select id from products where name = 'Retrait RIA'), 'amount', 10000, 'details', 'MTCN 1234567890'))));

-- Une sous-gérante à l'agence principale, et un fournisseur avec un achat à crédit.
select set_config('app.uid', '00000000-0000-0000-0000-000000000001', false);
insert into auth.users(id, email, raw_user_meta_data) values
  ('00000000-0000-0000-0000-000000000006', 'sousgerante@ayael.app', '{"full_name":"Sous-gérante (démo)"}')
on conflict do nothing;
select save_profile('{"id":"00000000-0000-0000-0000-000000000006","role":"sous_gerant","active":true,"perm_appro":true,"perm_fournisseurs":true}');
select save_partenaire('{"name":"Moov Distribution","kind":"fournisseur","country":"Togo","contacts":"90 00 00 00","products":"Solde électronique Moov Money, crédit Moov"}');
select op_approvisionnement(jsonb_build_object('gtp_id', (select id from accounts where name = 'Mixx agent 1'), 'amount', 100000,
  'source_id', (select id from accounts where name = 'Moov Distribution'), 'details', 'Solde électronique Mixx (démo)'));

-- Partenaire en Côte d'Ivoire : l'exemple du cahier des charges (il doit 400 000 F à AYAEL).
select save_partenaire('{"name":"Partenaire CI (démo)","kind":"les_deux","country":"Côte d''Ivoire","contacts":"M. Koné · +225 07 00 00 00",
  "products":"Cartes physiques, alimentation UBA / GTCO, recharges Ecobank et BNI, paiements de bénéficiaires, Wave / Orange Money / MTN",
  "conditions":"Frais Ecobank / BNI : 1 % du montant","payment_terms":"Compensation chaque fin de semaine","cartes_compensees":false}');
select op_ajustement(jsonb_build_object('account_id', (select id from accounts where name = 'Coffre'), 'target', 3000000, 'details', 'Coffre (démo)'));
select create_dossier(jsonb_build_object('partenaire_id', (select id from partenaires where name = 'Partenaire CI (démo)'), 'type', 'paiement',
  'beneficiary', 'KOFFI AMAVI', 'beneficiary_phone', '90 55 44 33', 'amount', 2000000, 'details', 'Demande du partenaire (démo)'));
select dossier_action(jsonb_build_object('dossier_id', (select max(id) from partner_ops), 'action', 'paiement',
  'account_id', (select id from accounts where name = 'Coffre')));
select create_dossier(jsonb_build_object('partenaire_id', (select id from partenaires where name = 'Partenaire CI (démo)'), 'type', 'appro',
  'lines', jsonb_build_array(jsonb_build_object('account_id', (select id from accounts where name = 'GTCO principal'), 'amount', 1500000),
                             jsonb_build_object('account_id', (select id from accounts where name = 'UBA AfricSika'), 'amount', 1000000))));
select dossier_action(jsonb_build_object('dossier_id', (select max(id) from partner_ops), 'action', 'alimentation',
  'account_id', (select id from accounts where name = 'GTCO principal'), 'amount', 1500000, 'reference', 'GTCO-DEMO-01'));
select create_dossier(jsonb_build_object('partenaire_id', (select id from partenaires where name = 'Partenaire CI (démo)'), 'type', 'recharge',
  'client_id', (select id from clients where name = 'DOPENYO AYAWA FLORENCE'), 'service', 'BNI', 'card_ref', '4521000012349316',
  'amount', 100000, 'billed', 102000, 'due_partner', 100000, 'encaisse', 102000));
select dossier_action(jsonb_build_object('dossier_id', (select max(id) from partner_ops), 'action', 'execution', 'reference', 'BNI-OK-77'));
-- Une recharge Ecobank encore en attente (payée par le client, pas encore faite par le partenaire).
select create_dossier(jsonb_build_object('partenaire_id', (select id from partenaires where name = 'Partenaire CI (démo)'), 'type', 'recharge',
  'client_id', (select id from clients where name = 'YASMINE DAOUDA'), 'service', 'Ecobank', 'card_ref', '5199000077771234',
  'amount', 50000, 'billed', 51000, 'encaisse', 51000));
-- Crédits clients : CHARIFE ISSAKA autorisé (plafond 150 000 F, 15 jours) ; l'exemple du cahier des charges.
select save_client_credit(jsonb_build_object('client_id', (select id from clients where name = 'CHARIFE ISSAKA'),
  'credit_allowed', true, 'credit_limit', 150000, 'credit_days', 15));
select op_vente(jsonb_build_object('client_id', (select id from clients where name = 'CHARIFE ISSAKA'), 'credit_amount', 50000,
  'lines', jsonb_build_array(jsonb_build_object('type', 'recharge', 'card_id', (select id from cards where card_number = '5027658'),
                             'gtp_id', (select id from accounts where name = 'GTCO principal'), 'amount', 50000)), 'force_duplicate', true));
update client_credits set created_at = now() - interval '20 days', due_date = current_date - 5 where id = (select max(id) from client_credits);
select op_vente(jsonb_build_object('client_id', (select id from clients where name = 'CHARIFE ISSAKA'), 'credit_amount', 30000,
  'lines', jsonb_build_array(jsonb_build_object('type', 'caisse', 'product_id', (select id from products where name = 'Crédit de communication'), 'amount', 30000, 'details', 'Moov'))));
select op_reglement(jsonb_build_object('client_id', (select id from clients where name = 'CHARIFE ISSAKA'), 'amount', 60000, 'reference', 'Espèces (démo)'));
-- Cartes physiques : deux types, une commande reçue en partie.
select save_card_type('{"name":"Carte UBA AfricSika","bank":"UBA","sale_price":5000}');
select save_card_type('{"name":"Carte GTCO Visa","bank":"GTCO","sale_price":6000}');
select create_dossier(jsonb_build_object('partenaire_id', (select id from partenaires where name = 'Partenaire CI (démo)'), 'type', 'cartes',
  'invoice_ref', 'FAC-CI-2026-031', 'lines', jsonb_build_array(
    jsonb_build_object('card_type_id', (select id from card_types where name = 'Carte UBA AfricSika'), 'qty', 50, 'unit_price', 3500),
    jsonb_build_object('card_type_id', (select id from card_types where name = 'Carte GTCO Visa'), 'qty', 20, 'unit_price', 4000))));
select dossier_action(jsonb_build_object('dossier_id', (select max(id) from partner_ops), 'action', 'reception',
  'lines', jsonb_build_array(jsonb_build_object('card_type_id', (select id from card_types where name = 'Carte UBA AfricSika'), 'qty', 30))));

select set_config('app.uid', '', false);
