// Importa as ferramentas necessárias para controlar o WhatsApp, ler Excel e mexer em pastas
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Configura o robô para manter a sessão ligada e abre o navegador visível para acompanhamento
const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        headless: true,
        protocolTimeout: 300000,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    }
});

// Define o número de telefone que vai receber o aviso de "Sucesso" ao final de cada grupo
const NUMERO_AVISO = "554799666043@c.us";

// Gera o código QR no terminal para você escanear com o celular
client.on('qr', (qr) => qrcode.generate(qr, { small: true }));

// Quando o WhatsApp termina de carregar, executa os comandos abaixo
client.on('ready', async () => {
    console.log('\n✅ WhatsApp Conectado!');

    try {
        // Abre o arquivo Excel chamado 'grupos.xlsm' que está na mesma pasta do robô
        const workbook = XLSX.readFile('./grupos.xlsm');
        // Seleciona a segunda aba da planilha (índice 1)
        const nomeDaAbaAlvo = 'GrupoW';
        const sheetName = workbook.SheetNames.find(n => n === nomeDaAbaAlvo);
        // Verificação de segurança: Se não achar o nome, ele avisa e para o robô
        if (!sheetName) {
            console.error(`❌ Erro: Não encontrei a aba chamada "${nomeDaAbaAlvo}" no Excel!`);
            fs.writeFileSync('./status.txt', 'ERRO');
            process.exit();
        }
        // Converte os dados da planilha em uma lista que o código consegue ler
        const dados = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        console.log(`📊 Total de linhas encontradas: ${dados.length}`);

        // Inicia a leitura de cada linha da sua planilha uma por uma
        for (const linha of dados) {
            // Procura automaticamente a coluna que tem o nome "grupo" ou "cliente"
            const chaves = Object.keys(linha);
            const chaveCerta = chaves.find(k => k.toLowerCase().includes("grupo") || k.toLowerCase().includes("cliente")) || chaves[0];

            // Pega o código (ex: 2025-001) e limpa espaços vazios
            const codigoBruto = linha[chaveCerta];
            const codigoExcel = codigoBruto ? String(codigoBruto).trim() : null;

            // Se a linha estiver vazia ou sem código válido, o robô pula para a próxima
            if (!codigoExcel || codigoExcel === "undefined" || codigoExcel === "") {
                console.log("⏭️ Linha sem código válido, pulando...");
                continue;
            }

            console.log(`\n🧐 Processando código: "${codigoExcel}"`);

            try {
                // Define o caminho da rede onde as pastas dos clientes estão guardadas
                const caminhoBase = 'Z:\\Diversos\\Processos\\Topografia\\Orçamentos\\2025';

                // Verifica se o computador consegue acessar a unidade Z: agora
                if (!fs.existsSync(caminhoBase)) {
                    console.error("❌ Erro: Unidade Z: não encontrada.");
                    break;
                }

                // Lê todos os nomes de arquivos e pastas dentro do caminho base
                const itensNoCaminho = fs.readdirSync(caminhoBase);
                // Procura a pasta que começa com o código do Excel (ex: começa com "2025-001")
                const nomePastaCompleto = itensNoCaminho.find(item =>
                    item.startsWith(codigoExcel) && fs.lstatSync(path.join(caminhoBase, item)).isDirectory()
                );

                // Se não encontrar nenhuma pasta com esse código, avisa e pula para o próximo cliente
                if (!nomePastaCompleto) {
                    console.log(`⚠️ Pasta não encontrada para: ${codigoExcel}`);
                    continue;
                }

                console.log(`📁 Pasta localizada: ${nomePastaCompleto}`);

                // Puxa a lista de todas as suas conversas atuais do WhatsApp
                const chats = await client.getChats();
                // Verifica se já existe um grupo criado com o nome exato da pasta
                const grupoExistente = chats.find(c => c.isGroup && c.name === nomePastaCompleto);

                let groupId;
                if (grupoExistente) {
                    // Se o grupo já existe, o robô apenas entra nele para enviar arquivos
                    console.log(`♻️ Grupo já existe.`);
                    groupId = grupoExistente.id._serialized;
                } else {
                    // Se o grupo não existe, cria um novo com os números listados abaixo
                    console.log(`🚀 Criando NOVO grupo...`);
                    const resultado = await client.createGroup(nomePastaCompleto, ["554799447847@c.us", "554799020133@c.us", "554796594112@c.us"]);
                    // Aguarda 8 segundos para o WhatsApp registrar a criação do grupo no servidor
                    await new Promise(r => setTimeout(r, 8000));
                    groupId = resultado.gid._serialized;
                }

                // Entra na pasta do cliente para ver quais arquivos tem lá dentro
                const caminhoFinal = path.join(caminhoBase, nomePastaCompleto);
                const arquivos = fs.readdirSync(caminhoFinal);

                // Filtra apenas os arquivos permitidos e ignora subpastas ou arquivos temporários
                const arquivosValidos = arquivos.filter(f => {
                    const ext = path.extname(f).toLowerCase();
                    const permitidos = ['.pdf', '.xlsx', '.xls', '.docx', '.jpg', '.png', '.jpeg', '.kml', '.xlsm'];
                    // Verifica se é um arquivo real (pula pastas amarelas) e se a extensão está na lista
                    return fs.lstatSync(path.join(caminhoFinal, f)).isFile() && permitidos.includes(ext) && !f.startsWith('~$');
                });

                // Se encontrar arquivos válidos, começa o envio
                if (arquivosValidos.length > 0) {
                    console.log(`📦 Enviando ${arquivosValidos.length} arquivos...`);
                    for (const arquivo of arquivosValidos) {
                        // Prepara o arquivo e envia para o grupo criado
                        const media = MessageMedia.fromFilePath(path.join(caminhoFinal, arquivo));
                        await client.sendMessage(groupId, media);
                        // Espera 2,5 segundos entre um arquivo e outro para evitar bloqueio
                        await new Promise(r => setTimeout(r, 2500));
                    }
                    console.log(`✅ Sucesso!`);
                    // Envia uma mensagem de confirmação para o seu número pessoal
                    await client.sendMessage(NUMERO_AVISO, `📢 *Robô:* Grupo *${nomePastaCompleto}* criado!`);
                }

            } catch (err) {
                // Caso aconteça um erro com um cliente específico, avisa no terminal e vai para o próximo
                console.error(`❌ Erro no cliente ${codigoExcel}:`, err.message);
            }
        }

        // Após processar toda a planilha, avisa que terminou
        console.log('\n🏁 TUDO PRONTO!');
        // CRIA O ARQUIVO QUE O EXCEL ESTÁ ESPERANDO PARA DAR O AVISO FINAL
        fs.writeFileSync('./status.txt', 'CONCLUIDO');

    } catch (err) {
        // Se der um erro grave (como planilha aberta ou sem rede), avisa o erro
        console.error('❌ Erro Crítico:', err.message);
        fs.writeFileSync('./status.txt', 'ERRO');
    }

    // Espera 5 segundos e encerra o navegador e o processo do robô
    console.log('Finalizando processo...');
    await new Promise(r => setTimeout(r, 5000));
    await client.destroy();
    process.exit();
});

// Comando que inicia oficialmente a conexão com o WhatsApp
client.initialize();
