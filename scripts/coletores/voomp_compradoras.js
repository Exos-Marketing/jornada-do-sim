async () => {
  // Coletor de compradoras da Voomp Creators (skill jornada-do-sim). Só faz GET.
  // Igual ao voomp.js da acompanhar-vendas, mas guarda o contato de quem pagou: é para a lista de ação.
  // O token é lido do cookie dentro da página e não sai dela.
  const CFG = __CFG__;
  const saida = { plataforma: 'voomp', coletado_em: new Date().toISOString() };
  if (location.hostname !== 'adm.voompcreators.com.br') return { ...saida, erro: 'sem_sessao', url: location.href };
  const api = 'https://apiadm.voompcreators.com.br/api';

  let cabecalhos = null;
  for (let tentativa = 0; tentativa < 6 && !cabecalhos; tentativa++) {
    const bruto = document.cookie.split('; ').find(c => c.startsWith('access_token='));
    if (bruto) {
      const token = decodeURIComponent(bruto.slice('access_token='.length));
      const tentar = { Authorization: /^bearer /i.test(token) ? token : 'Bearer ' + token, Accept: 'application/json' };
      if ((await fetch(api + '/login/get-user', { headers: tentar })).ok) cabecalhos = tentar;
    }
    if (!cabecalhos) await new Promise(resolver => setTimeout(resolver, 2500));
  }
  if (!cabecalhos) return { ...saida, erro: 'sem_sessao', url: location.href };

  const hash = async (texto) => {
    if (!texto) return null;
    const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(String(texto).trim().toLowerCase()));
    return [...new Uint8Array(bytes)].slice(0, 6).map(b => b.toString(16).padStart(2, '0')).join('');
  };

  const filtro = 'type=created_at&per_page=100'
    + '&date_start=' + encodeURIComponent(CFG.desde + 'T03:00:00.000Z')
    + '&date_end=' + encodeURIComponent(new Date(Date.now() + 864e5).toISOString())
    + CFG.produtos.map(p => '&products[]=' + p).join('');

  const linhas = [];
  let paginas = 1;
  for (let pagina = 1; pagina <= paginas && pagina <= 300; pagina++) {
    const resposta = await fetch(api + '/sale?' + filtro + '&page=' + pagina, { headers: cabecalhos });
    if (resposta.status === 401 || resposta.status === 403) return { ...saida, erro: 'sem_sessao', http: resposta.status };
    if (!resposta.ok) return { ...saida, erro: 'http_' + resposta.status, pagina };
    const dados = (await resposta.json()).data;
    paginas = dados.last_page;
    for (const v of dados.data) {
      const cliente = v.client || {};
      const pago = v.status === 'paid';
      linhas.push({
        id: v.id,
        produto_id: v.product_id,
        status: v.status,
        criado: v.created_at,
        pago: v.paid_at,
        reembolsado: v.refunded_at,
        h_doc: await hash(String(cliente.document || '').replace(/\D/g, '')),
        h_email: await hash(cliente.email),
        nome: pago ? cliente.name : null,
        telefone: pago ? cliente.cellphone : null,
        email: pago ? cliente.email : null,
      });
    }
  }
  return { ...saida, linhas };
}
