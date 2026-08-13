// Estrutura em memória para armazenar os produtos lidos do Excel
let produtosDB = {};
let graficoInstancia = null;

// Formatação monetária padronizada
function formatarMoeda(valor) {
  return Number(valor).toLocaleString('pt-BR', {
    style: 'currency',
    currency: 'BRL'
  });
}

/**
 * Lê o ficheiro Excel selecionado pelo usuário utilizando a biblioteca SheetJS
 */
function carregarExcel(event) {
  const file = event.target.files[0];
  const statusLabel = document.getElementById('fileStatus');
  
  if (!file) {
    statusLabel.innerText = "Nenhum ficheiro carregado";
    return;
  }

  statusLabel.innerText = `Lendo "${file.name}"...`;

  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // Obtém a primeira folha da pasta de trabalho
      const primeiraAba = workbook.SheetNames[0];
      const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[primeiraAba]);

      if (!linhas || linhas.length === 0) {
        alert("O ficheiro Excel está vazio ou não possui cabeçalhos válidos.");
        statusLabel.innerText = "Ficheiro inválido";
        return;
      }

      // Agrupar e estruturar os dados por produto
      produtosDB = {};

      linhas.forEach(linha => {
        // Suporta diferentes convenções de cabeçalhos
        const id = String(linha.id_produto || linha.ID || linha.id || '').trim();
        const nome = linha.nome_produto || linha.Nome || linha.produto || id;
        const dataRegisto = String(linha.data || linha.Data || '');
        const preco = parseFloat(linha.preco || linha.Preco || linha.valor || 0);
        const desconto = linha.desconto_anunciado || linha.desconto || '';

        if (!id) return;

        if (!produtosDB[id]) {
          produtosDB[id] = {
            id: id,
            nome: nome,
            datas: [],
            precos: [],
            descontoAnunciado: desconto
          };
        }

        produtosDB[id].datas.push(dataRegisto);
        produtosDB[id].precos.push(preco);
        if (desconto) {
          produtosDB[id].descontoAnunciado = desconto;
        }
      });

      const totalProdutos = Object.keys(produtosDB).length;

      if (totalProdutos === 0) {
        alert("Nenhum dado válido de produtos foi identificado. Verifique os nomes das colunas.");
        statusLabel.innerText = "Estrutura incorreta";
        return;
      }

      popularSeletor();
      statusLabel.innerText = `✅ Sucesso: ${totalProdutos} produto(s) carregado(s)!`;

    } catch (error) {
      console.error("Erro ao processar o Excel:", error);
      alert("Ocorreu um erro ao processar o arquivo Excel.");
      statusLabel.innerText = "Erro ao processar ficheiro";
    }
  };

  reader.readAsArrayBuffer(file);
}

/**
 * Preenche o elemento <select> com os produtos extraídos do Excel
 */
function popularSeletor() {
  const select = document.getElementById('productSelect');
  select.innerHTML = '<option value="">Selecione um produto da lista...</option>';

  Object.keys(produtosDB).forEach(id => {
    const opcao = document.createElement('option');
    opcao.value = id;
    opcao.textContent = produtosDB[id].nome;
    select.appendChild(opcao);
  });

  select.disabled = false;
  document.getElementById('btnAnalisar').disabled = false;
}

/**
 * Analisa as variações históricas e emite o diagnóstico
 */
function analisarPreco() {
  const select = document.getElementById('productSelect');
  const id = select.value;

  if (!id || !produtosDB[id]) {
    alert("Por favor, selecione um produto para analisar.");
    return;
  }

  const produto = produtosDB[id];
  const historicoPrecos = produto.precos;

  if (historicoPrecos.length === 0) return;

  // O último elemento corresponde ao preço atual
  const precoAtual = historicoPrecos[historicoPrecos.length - 1];
  const precosPassados = historicoPrecos.length > 1 ? historicoPrecos.slice(0, -1) : historicoPrecos;

  // Cálculos Estatísticos
  const mediaHistorica = precosPassados.reduce((total, valor) => total + valor, 0) / precosPassados.length;
  const menorPreco = Math.min(...historicoPrecos);
  const precoImediatamenteAnterior = precosPassados[precosPassados.length - 1];

  // Atualização dos elementos da interface
  document.getElementById('resultCard').style.display = 'block';
  document.getElementById('productTitle').innerText = produto.nome;
  document.getElementById('productBadgeDesconto').innerText = produto.descontoAnunciado || 'OFERTA';
  document.getElementById('currentPrice').innerText = formatarMoeda(precoAtual);
  document.getElementById('avgPrice').innerText = formatarMoeda(mediaHistorica);
  document.getElementById('minPrice').innerText = formatarMoeda(menorPreco);

  // Elementos do veredito
  const verdictBox = document.getElementById('verdictBox');
  const verdictTitle = document.getElementById('verdictTitle');
  const verdictDesc = document.getElementById('verdictDesc');

  verdictBox.className = 'verdict-box';

  // Lógica de Detecção do "Metade do Dobro"
  if (precoAtual <= menorPreco) {
    // Menor preço de todo o histórico
    verdictBox.classList.add('badge-bom');
    verdictTitle.innerText = '🟢 Desconto Real!';
    verdictDesc.innerText = `Este é o menor valor registrado em todo o histórico. A promoção anunciada é autêntica e vantajosa.`;
  } else if (precoImediatamenteAnterior > mediaHistorica * 1.15 && precoAtual >= mediaHistorica * 0.95) {
    // O preço subiu consideravelmente pouco antes da suposta "oferta"
    verdictBox.classList.add('badge-alerta');
    verdictTitle.innerText = '🔴 Alerta: "Metade do Dobro"!';
    verdictDesc.innerText = `O preço foi aumentado repentinamente antes desta promoção. O valor atual (${formatarMoeda(precoAtual)}) está apenas no nível normal histórico.`;
  } else {
    // Variação comum de mercado
    verdictBox.classList.add('badge-neutro');
    verdictTitle.innerText = '🟡 Promoção Neutra / Comum';
    verdictDesc.innerText = `O preço está dentro da margem média de oscilação dos últimos períodos. Não há uma economia fora do padrão.`;
  }

  // Gera ou atualiza o gráfico
  desenharGrafico(produto.datas, produto.precos);
}

/**
 * Criação e renderização do gráfico com Chart.js
 */
function desenharGrafico(labels, precos) {
  const ctx = document.getElementById('priceChart').getContext('2d');

  if (graficoInstancia) {
    graficoInstancia.destroy();
  }

  graficoInstancia = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: 'Histórico de Preço',
        data: precos,
        borderColor: '#2563eb',
        backgroundColor: 'rgba(37, 99, 235, 0.08)',
        borderWidth: 2.5,
        fill: true,
        tension: 0.3,
        pointBackgroundColor: '#2563eb',
        pointRadius: 5,
        pointHoverRadius: 7
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(context) {
              return ' Valor: ' + formatarMoeda(context.parsed.y);
            }
          }
        }
      },
      scales: {
        y: {
          ticks: {
            callback: function(val) {
              return formatarMoeda(val);
            }
          },
          grid: {
            color: '#e2e8f0'
          }
        },
        x: {
          grid: {
            display: false
          }
        }
      }
    }
  });
}