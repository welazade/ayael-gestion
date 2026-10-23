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
update settings set value = value || '{"address":"Lomé, Togo (adresse à compléter)","phone":"+228 00 00 00 00"}' where key = 'company';

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

select set_config('app.uid', '', false);
