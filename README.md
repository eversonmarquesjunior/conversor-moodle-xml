# Conversor Moodle XML

Ferramenta web para converter arquivos de questões em **Word (.docx)** para o formato **XML do Moodle**, pronto para importar em um banco de questões.

Acesse em: https://eversonmarquesjunior.github.io/conversor-moodle-xml/

## O que o site faz

1. **Upload**: o usuário arrasta (em qualquer ponto da página) ou seleciona um ou vários arquivos `.docx` contendo as questões.
2. **Leitura do Word**: cada questão deve estar organizada em uma tabela no documento, com linhas identificadas pelos rótulos `ENUNCIADO`, `CORRETA` e `Incorreta` na primeira coluna.
3. **Conversão**: o site lê o XML interno do `.docx` diretamente no navegador (sem enviar o arquivo para nenhum servidor) e extrai:
   - Texto do enunciado e das alternativas, preservando **negrito**, quebras de linha e **listas** (numeradas, romanas e com marcadores).
   - **Imagens** embutidas nas questões, exportadas como arquivos anexos ao XML no formato que o Moodle espera.
   - **Links de referência** (URLs), convertidos em hyperlinks clicáveis.
4. **Download**: gera o(s) arquivo(s) `.xml` no padrão de importação do Moodle (um único `.xml` para um arquivo, ou um `.zip` quando vários `.docx` são convertidos de uma vez).

## Como funciona por dentro

- **100% front-end**: todo o processamento acontece no navegador do usuário, em JavaScript puro. Não há backend nem envio de dados para fora.
- **Bibliotecas usadas**: [JSZip](https://stuk.github.io/jszip/) para ler o `.docx` (que é um arquivo `.zip`) e montar o `.zip` de saída quando necessário.
- **Arquivos do projeto**:
  - `index.html` — estrutura da página
  - `style.css` — visual (tema escuro)
  - `script.js` — toda a lógica de leitura do `.docx` e geração do XML

## Formato esperado do arquivo Word

Cada questão deve ser uma tabela com uma linha por elemento:

| Rótulo (1ª coluna) | Conteúdo |
|---|---|
| `ENUNCIADO` | Texto (e imagens) da pergunta |
| `CORRETA` | Alternativa correta |
| `Incorreta` | Alternativa incorreta (uma linha por alternativa errada) |

## Limitações conhecidas

- Fórmulas matemáticas inseridas como **Equação do Word** (Cálculo, frações, integrais, etc.) ainda não são convertidas — atualmente precisam ser inseridas como imagem/print no `.docx`.
- Questões cuja tabela não segue os rótulos esperados podem não ser identificadas na conversão.
