# FamaTur · Dashboard de Adesão ao Aplicativo

Dashboard estático para importar:
1. planilha de funcionários;
2. logs de utilização do aplicativo de um dia.

## Regra de negócio

A base considera somente estes cargos:
- MOTORISTA - ÔNIBUS
- MOTORISTA - MICRO ÔNIBUS
- MOTORISTA - VAN

Um motorista é considerado **"Usando"** quando seu nome aparece pelo menos uma vez na planilha de logs daquele dia.

Os logs que pertencem a outras funções (ex.: fiscal) são ignorados.

## Histórico

O histórico é salvo no `localStorage` do navegador. Isso permite importar um arquivo por dia e acompanhar:
- adesão diária;
- total de motoristas;
- usando x sem uso;
- quantidade de eventos;
- histórico geral;
- evolução individual de cada motorista.

**Importante:** nesta versão, o histórico é local ao navegador/dispositivo. Se você abrir o site em outro computador, ele não terá o histórico deste computador.

Para um histórico centralizado entre computadores/usuários, a próxima etapa é ligar o projeto a Supabase/PostgreSQL.

## Rodar localmente

Pode abrir `index.html` diretamente no navegador. Para desenvolvimento, também pode usar qualquer servidor estático.

## Publicar na Vercel

1. Crie um repositório no GitHub.
2. Envie `index.html`, `style.css`, `app.js` e `README.md`.
3. Na Vercel, importe o repositório.
4. Framework Preset: **Other**.
5. Build Command: deixe vazio.
6. Output Directory: deixe vazio.
7. Deploy.

Não há backend obrigatório nesta primeira versão.

## Formato esperado

### Funcionários
A aplicação procura:
- `NOME`
- `CARGO`
- `EMPRESA`
- `MATRÍCULA` (opcional)

### Logs
A aplicação procura:
- `DATA/HORA`
- `MOTORISTA`
- `AÇÃO` (opcional)
- `TELA` (opcional)

A aplicação também normaliza acentos e espaços para evitar erro simples de correspondência de nomes.
