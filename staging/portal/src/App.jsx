import { useEffect, useState, useRef } from 'react';
import {
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut
} from 'firebase/auth';
import { collection, getDocs, getDoc, getDocsFromServer, getDocFromServer, doc, setDoc, updateDoc, serverTimestamp, writeBatch } from 'firebase/firestore';
import { auth, db } from './firebase';
import './App.css';
import pilotSource from '../../verde-levato-pilot.mjs?raw';
import { buildVerdeFrame } from './verde-frame.mjs';
import { loadTenantVerde } from '../../verde-tenant-loader.mjs';
const VERDE_SERVER_ENABLED = false; // Enable only after staging rules and browser verification.
import { resolveTenantContext } from '../../tenant-context.mjs';


function TenantPortal({ user, companies }) {
  const [selected, setSelected] = useState(null);
  const [verdeOpen, setVerdeOpen] = useState(false);
  const verdeRef = useRef(null);
  const [verdeFrame, setVerdeFrame] = useState('');
  const [verdeError, setVerdeError] = useState('');
  const revision = useRef(0);
  function closeVerde() { revision.current++; setVerdeOpen(false); setVerdeFrame(''); setVerdeError(''); }
  async function openVerde() {
    const request = ++revision.current;
    setVerdeError(''); setVerdeOpen(false); setVerdeFrame('');
    try {
      const records = VERDE_SERVER_ENABLED ? await loadTenantVerde({db,user,tenantId:selected.id,sdk:{doc,collection,getDocFromServer,getDocsFromServer}}) : [];
      if (request !== revision.current || auth.currentUser?.uid !== user.uid) return;
      setVerdeFrame(buildVerdeFrame(pilotSource,records)); setVerdeOpen(true);
    } catch { if (request === revision.current) setVerdeError('Archivio non disponibile.'); }
  }
  useEffect(() => () => { revision.current++; }, []);
  useEffect(() => {
    const close = event => {
      if (event.origin === 'null' && event.source === verdeRef.current?.contentWindow && event.data === 'varga-verde-close') closeVerde();
    };
    window.addEventListener('message', close);
    return () => window.removeEventListener('message', close);
  }, []);

  const roles = {
    admin: 'Amministratore',
    supervisor: 'Caposquadra',
    employee: 'Dipendente'
  };

  return (
    <main className="dashboard">
      <header>
        <div>
          <h2>VARGA CANTIERI 2.0</h2>
          <small>Portale aziendale - STAGING</small>
        </div>
        <button onClick={() => signOut(auth)}>Esci</button>
      </header>

      <section className="content">
        <h1>Benvenuto</h1>
        <p>{user.displayName || user.email}</p>

        {!selected ? (
          <>
            <h2>Seleziona la tua azienda</h2>

            {companies.length === 0 && (
              <p>Nessuna azienda attiva disponibile.</p>
            )}

            {companies.map(company => (
              <article className="company" key={company.id}>
                <div>
                  <h3>{company.name}</h3>
                  <small>{roles[company.role]}</small>
                </div>
                <button onClick={() => setSelected(company)}>
                  Accedi
                </button>
              </article>
            ))}
          </>
        ) : (
          <>
            <button onClick={() => { closeVerde(); setSelected(null); }}>
              Cambia azienda
            </button>

            <h2>{selected.name}</h2>
            <p>Ruolo: {roles[selected.role]}</p>

            <article className="company">
              <div>
                <h3>Area aziendale</h3>
                <p>Accesso effettuato correttamente.</p>
                <small>Moduli operativi in sviluppo.</small>
                <p><button onClick={openVerde}>🌱 Verde Levato — anteprima demo</button></p>
              </div>
            </article>
            {verdeError && <p role="alert">{verdeError}</p>}
            {verdeOpen && (
              <section aria-label="Anteprima Verde Levato">
                <button onClick={closeVerde}>Chiudi Verde Levato</button>
                <p>Archivio dimostrativo vuoto. Nessun collegamento ai dati aziendali reali.</p>
                <iframe key={selected.id} title="Verde Levato demo" ref={verdeRef} srcDoc={verdeFrame} sandbox="allow-scripts" style={{width:'100%',height:'75vh',border:'1px solid #ccd5ce',borderRadius:8,background:'white'}} />
              </section>
            )}
          </>
        )}
      </section>
    </main>
  );
}

export default function App() {
  const [user, setUser] = useState(null);
  const [allowed, setAllowed] = useState(false);
  const [isSuper, setIsSuper] = useState(false);
  const [myCompanies, setMyCompanies] = useState([]);

  const [loading, setLoading] = useState(true);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [tenants, setTenants] = useState([]);
  const [busy, setBusy] = useState(false);
  const [selectedCompany, setSelectedCompany] = useState(null);
  const [members, setMembers] = useState([]);
  const [membersLoading, setMembersLoading] = useState(false);
  const [membersError, setMembersError] = useState('');


  useEffect(() => {
    let revision = 0;
    const unsubscribe = onAuthStateChanged(auth, async current => {
      const id = ++revision;
      setLoading(true);
      setAllowed(false);
      setIsSuper(false);
      setMyCompanies([]);

      setTenants([]);
      setSelectedCompany(null);
      setMembers([]);

      setUser(current);

      try {
        if (current) {
          const token = await current.getIdTokenResult(true);
          if (id !== revision) return;


          if (
            token.claims.email_verified !== true ||
            token.claims.apple_review === true
          ) {
            setError('Account non autorizzato.');
            await signOut(auth);
            return;
          }

          if (token.claims.superAdmin !== true) {
            const memberships = await getDocs(
              collection(db, 'users', current.uid, 'memberships')
            );

            const companies = [];

            for (const membership of memberships.docs) {
              if (membership.data().status !== 'active') continue;

              try {
                const tenant = await getDoc(
                  doc(db, 'tenants', membership.id)
                );

                const member = await getDoc(
                  doc(db, 'tenants', membership.id,
                      'members', current.uid)
                );

                const context = resolveTenantContext({
                  uid: current.uid,
                  claims: token.claims,
                  tenantId: membership.id,
                  tenant: tenant.exists() ? tenant.data() : null,
                  member: member.exists() ? member.data() : null,
                  membership: membership.data()
                });
                if (context) {
                  companies.push({
                    id: context.tenantId,
                    name: tenant.data().name,
                    role: context.role
                  });
                }
              } catch (e) {
                if (e.code !== 'permission-denied') throw e;
              }
            }

            if (id !== revision) return;

            setMyCompanies(companies);
            setAllowed(true);
            return;
          }

          setIsSuper(true);

          const snapshot = await getDocs(collection(db, 'tenants'));
          if (id !== revision) return;

          setTenants(snapshot.docs.map(d => ({
            ...d.data(),
            id: d.id
          })));
          setAllowed(true);
        }
      } catch {
        if (id === revision) {
          setError('Impossibile verificare accesso o aziende.');
        }
      } finally {
        if (id === revision) setLoading(false);
      }
    });

    return () => {
      revision++;
      unsubscribe();
    };
  }, []);




  async function updateMember(member, changes) {
    if (!allowed || busy || membersLoading ||
        !selectedCompany || !user ||
        member.id === user.uid ||
        member.demo !== true) return;

    const company = selectedCompany;

    setBusy(true);
    setMembersError('');

    try {
      const batch = writeBatch(db);

      batch.update(
        doc(db, 'tenants', company.id, 'members', member.id),
        changes
      );

      batch.update(
        doc(db, 'users', member.id, 'memberships', company.id),
        changes
      );

      await batch.commit();
      await showMembers(company);

    } catch (e) {
      setMembersError(
        'Modifica non riuscita: ' + (e.code || 'errore')
      );
    } finally {
      setBusy(false);
    }
  }

  async function changeMemberRole(member, role) {
    if (!['admin', 'supervisor', 'employee'].includes(role)) return;
    if (role === member.role) return;

    if (!window.confirm(
      'Confermi la modifica del ruolo di ' +
      (member.displayName || 'questo utente') + '?'
    )) return;

    await updateMember(member, { role });
  }

  async function toggleMember(member) {
    const status = member.status === 'active'
      ? 'suspended' : 'active';

    if (!window.confirm(
      status === 'suspended'
        ? 'Sospendere questa appartenenza?'
        : 'Riattivare questa appartenenza?'
    )) return;

    await updateMember(member, { status });
  }

  async function showMembers(company) {
    setSelectedCompany(company);
    setMembers([]);
    setMembersError('');
    setMembersLoading(true);

    try {
      const snapshot = await getDocs(
        collection(db, 'tenants', company.id, 'members')
      );

      setMembers(snapshot.docs.map(d => ({
        ...d.data(),
        id: d.id
      })));
    } catch (e) {
      setMembersError('Impossibile caricare gli utenti.');
    } finally {
      setMembersLoading(false);
    }
  }

  async function refreshCompanies() {
    const snapshot = await getDocs(collection(db, 'tenants'));
    setTenants(snapshot.docs.map(d => ({
      ...d.data(),
      id: d.id
    })));
  }

  async function companyOperation(operation) {
    if (busy || !allowed) return;
    setBusy(true);
    setError('');

    try {
      await operation();
      await refreshCompanies();
    } catch (e) {
      setError('Operazione non riuscita: ' + e.code);
    } finally {
      setBusy(false);
    }
  }

  async function createCompany() {
    const name = window.prompt('Nome della nuova azienda demo:');
    if (!name || name.trim().length < 2 ||
        name.trim().length > 120) return;

    const id = 'demo-' + crypto.randomUUID();

    await companyOperation(() =>
      setDoc(doc(db, 'tenants', id), {
        name: name.trim(),
        status: 'active',
        demo: true,
        createdAt: serverTimestamp()
      })
    );
  }

  async function renameCompany(company) {
    const name = window.prompt(
      'Modifica nome azienda:', company.name
    );

    if (!name || name.trim().length < 2 ||
        name.trim().length > 120) return;

    await companyOperation(() =>
      updateDoc(doc(db, 'tenants', company.id), {
        name: name.trim()
      })
    );
  }

  async function toggleCompany(company) {
    const next = company.status === 'active'
      ? 'suspended' : 'active';

    const action = next === 'suspended'
      ? 'Sospendere' : 'Riattivare';

    if (!window.confirm(action + ' ' + company.name + '?')) return;

    await companyOperation(() =>
      updateDoc(doc(db, 'tenants', company.id), {
        status: next
      })
    );
  }

  async function login(event) {
    event.preventDefault();
    setError('');
    setBusy(true);

    try {
      await signInWithEmailAndPassword(auth, email, password);
    } catch {
      setError('Accesso non riuscito. Verifica le credenziali.');
    } finally {
      setPassword('');
      setBusy(false);
    }
  }

  if (loading) {
    return <main className="login">Caricamento...</main>;
  }

  if (user && !allowed) {
    return (
      <main className="login">
        <p>{error || 'Accesso non autorizzato'}</p>
        <button onClick={() => signOut(auth)}>Esci</button>
      </main>
    );
  }

  if (!user) {
    return (
      <main className="login">
        <form className="login-card" onSubmit={login}>
          <div className="logo">VC</div>
          <h1>VARGA CANTIERI 2.0</h1>
          <p>Portale Super Admin</p>
          <label>Indirizzo email</label>
          <input type="email" value={email}
            onChange={e => setEmail(e.target.value)}
            autoComplete="username" required />
          <label>Password</label>
          <input type="password" value={password}
            onChange={e => setPassword(e.target.value)}
            autoComplete="current-password" required />
          {error && <p className="error">{error}</p>}
          <button disabled={busy}>
            {busy ? 'Accesso in corso...' : 'ACCEDI'}
          </button>
          <small>STAGING - Ambiente di sviluppo</small>
        </form>
      </main>
    );
  }


  if (user && allowed && !isSuper) {
    return <TenantPortal
      key={user.uid}
      user={user}
      companies={myCompanies}
    />;
  }

  return (
    <main className="dashboard">
      <header>
        <div>
          <h2>VARGA CANTIERI 2.0</h2>
          <small>Super Admin - STAGING</small>
        </div>
        <button onClick={() => signOut(auth)}>Esci</button>
      </header>

      <section className="content">
        <h1>Dashboard amministrativa</h1>
        <p>{user.email}</p>

        <div className="stats">
          <article>
            <span>Aziende registrate</span>
            <strong>{tenants.length}</strong>
          </article>
          <article>
            <span>Aziende attive</span>
            <strong>
              {tenants.filter(t => t.status === 'active').length}
            </strong>
          </article>
        </div>

        
        <div className="section-heading">
          <h2>Gestione aziende</h2>
          <button disabled={busy} onClick={createCompany}>
            + Nuova azienda
          </button>
        </div>

        {error && <p className="dashboard-error">{error}</p>}


        {tenants.map(t => (
          <article className="company" key={t.id}>
            <div>
              <h3>{t.name}</h3>
              <small>{t.id}</small>
            </div>
            
            <div className="company-actions">
              <button
                disabled={busy || membersLoading}
                onClick={() => showMembers(t)}>
                Utenti
              </button>

              <span className={
                t.status === 'active'
                  ? 'status' : 'status-suspended'
              }>
                {t.status === 'active' ? 'Attiva' : 'Sospesa'}
              </span>

              {t.demo === true && (
                <>
                  <button className="edit-button"
                    disabled={busy}
                    onClick={() => renameCompany(t)}>
                    Modifica
                  </button>

                  <button
                    disabled={busy}
                    className={t.status === 'active'
                      ? 'suspend-button' : 'activate-button'}
                    onClick={() => toggleCompany(t)}>
                    {t.status === 'active'
                      ? 'Sospendi' : 'Riattiva'}
                  </button>
                </>
              )}
            </div>

          </article>
        ))}
        
        {selectedCompany && (
          <section className="members-panel">
            <div className="section-heading">
              <div>
                <h2>Utenti aziendali</h2>
                <p>{selectedCompany.name}</p>
              </div>
              <button
                disabled={membersLoading || busy}
                onClick={() => {
                  setSelectedCompany(null);
                  setMembers([]);
                }}>
                Chiudi
              </button>
            </div>

            {membersLoading && <p>Caricamento utenti...</p>}
            {membersError && <p>{membersError}</p>}

            {!membersLoading && !membersError &&
              members.length === 0 && (
                <p>Nessun utente assegnato.</p>
            )}

            {members.map(m => (
              <article className="member-card" key={m.id}>
                <div>
                  <h3>{m.displayName || 'Utente aziendale'}</h3>
                  <small>{m.email || m.id}</small>
                </div>
                <div>
                  <strong>
                    {{
                      admin: 'Amministratore',
                      supervisor: 'Caposquadra',
                      employee: 'Dipendente'
                    }[m.role] || m.role}
                  </strong>

                  {m.id === user.uid ? (
                    <p>Profilo protetto</p>
                  ) : m.demo === true && (
                    <div className="member-controls">
                      <label>Ruolo aziendale</label>

                      <select
                        value={m.role}
                        disabled={busy || membersLoading}
                        onChange={e =>
                          changeMemberRole(m, e.target.value)
                        }>
                        <option value="admin">
                          Amministratore
                        </option>
                        <option value="supervisor">
                          Caposquadra
                        </option>
                        <option value="employee">
                          Dipendente
                        </option>
                      </select>

                      <button
                        disabled={busy || membersLoading}
                        className={m.status === 'active'
                          ? 'suspend-button'
                          : 'activate-button'}
                        onClick={() => toggleMember(m)}>
                        {m.status === 'active'
                          ? 'Sospendi'
                          : 'Riattiva'}
                      </button>
                    </div>
                  )}

                  <p>
                    Appartenenza: {
                      m.status === 'active'
                        ? 'Attiva' : 'Sospesa'
                    }
                  </p>
                </div>
              </article>
            ))}

            <small>
              Gestione appartenenze utenti - Ambiente STAGING
            </small>
          </section>
        )}

<p className="notice">
          Gestione aziende abilitata. Ambiente dimostrativo.
        </p>
      </section>
    </main>
  );
}
