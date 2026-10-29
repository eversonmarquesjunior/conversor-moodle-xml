# Conversor Moodle XML

Ferramenta web para converter arquivos de questões em **Word (.docx)** para o formato **XML do Moodle**, pronto para importar em um banco de questões.

Acesse em: https://eversonmarquesjunior.github.io/conversor-moodle-xml/

## O que o site faz

1. **Upload**: o usuário arrasta (em qualquer ponto da página) ou seleciona um ou vários arquivos `.docx` contendo as questões.
2. **Leitura do Word**: cada questão deve estar organizada em uma tabela no documento, com linhas identificadas pelos rótulos `ENUNCIADO`, `CORRETA` e `Incorreta` na primeira coluna.
3. **Conversão**: o site lê o XML interno do `.docx` diretamente no navegador (sem enviar o arquivo para nenhum servidor) e extrai:
   - Texto do enunciado e das alternativas, preservando **negrito**, *itálico*, sublinhado, sobrescrito, subscrito, quebras de linha e **listas** do Word (números, letras, romanos e marcadores, em qualquer nível).
   - **Espaçamento do Word**: alinhamento, entrelinha, recuo e espaço entre parágrafos.
   - **Equações do Word** (frações, raízes, somatórios, integrais, matrizes…), convertidas em imagem PNG igual à do Word, com o texto da equação como descrição. Símbolos da fonte Symbol (α, Δ, ≤…) viram caracteres normais.
   - **Tabelas** dentro do enunciado, com células mescladas e cor de fundo.
   - **Imagens** embutidas nas questões, no tamanho do Word e centralizadas junto com o título (Figura…) e a fonte (Fonte: …), exportadas como arquivos anexos ao XML.
   - **Links de referência** (URLs), convertidos em hyperlinks clicáveis.
4. **Download**: gera o(s) arquivo(s) `.xml` no padrão de importação do Moodle (um único `.xml` para um arquivo, ou um `.zip` quando vários `.docx` são convertidos de uma vez).
5. **Relatório de curadoria** logo abaixo do botão, por arquivo:
   - **Fidelidade**: confere, questão por questão, se o texto, as imagens, as tabelas e as equações do Word chegaram ao XML.
   - **Estrutura**: número de alternativas fora do padrão do banco, nenhuma ou mais de uma correta, alternativa vazia, letra digitada e correta bem mais longa que as outras.
   - **Duplicidade**: alternativas com texto exatamente igual e questões com enunciado exatamente igual.
   - **Documento**: alterações controladas não aceitas, comentários, texto colorido, imagens EMF/WMF, equações antigas (Equation 3.0) e numeração das questões pulada ou repetida.
   - Mostra como cada questão fica no Moodle e permite copiar o relatório em tabela.

## Como funciona por dentro

- **100% front-end**: todo o processamento acontece no navegador do usuário, em JavaScript puro. Não há backend nem envio de dados para fora.
- **Bibliotecas usadas**: [JSZip](https://stuk.github.io/jszip/) para ler o `.docx` (que é um arquivo `.zip`) e montar o `.zip` de saída quando necessário; [MathJax](https://www.mathjax.org/), carregado só quando o documento tem equações, para desenhá-las.
- **Arquivos do projeto**:
  - `index.html` — estrutura da página
  - `style.css` — visual (tema escuro)
  - `script.js` — fluxo principal: leitura das questões do `.docx` e geração do XML
  - `formatacao-docx.js` — formatação do Word (estilos, listas, parágrafos, tabelas e espaçamento)
  - `equacoes.js` — equações do Word e símbolos (OMML → LaTeX → PNG)
  - `relatorio.js` — verificações e exibição do relatório de curadoria
- Parte da leitura do Word e das verificações foi adaptada do projeto *Curadoria Word × Moodle*.

## Formato esperado do arquivo Word

Cada questão deve ser uma tabela com uma linha por elemento:

| Rótulo (1ª coluna) | Conteúdo |
|---|---|
| `ENUNCIADO` | Texto (e imagens) da pergunta |
| `CORRETA` | Alternativa correta |
| `Incorreta` | Alternativa incorreta (uma linha por alternativa errada) |

## Limitações conhecidas

- Equações antigas (**Equation 3.0** / objetos incorporados) não são convertidas — refaça no editor de equações do Word (Inserir → Equação). O relatório aponta onde estão.
- Para desenhar as equações é preciso internet (o MathJax é carregado de CDN); sem ela, a equação vai como texto formatado e o relatório avisa.
- Imagens em EMF/WMF não aparecem no Moodle e ficam de fora; o relatório aponta quais são.
- Questões cuja tabela não segue os rótulos esperados podem não ser identificadas na conversão.
