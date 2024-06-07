# Importação da Folha de Pagamento

Este repositório contém ferramentas para importar a folha de pagamento da empresa.

## Passo a Passo

1. **Recebimento:** Receba a folha de pagamento por e-mail ou WhatsApp.
2. **Preparação:** 
   - Separe o mês atual em um novo arquivo Excel.
   - Salve o arquivo no formato `Folha (mês) (aaaa-mm-dd)`.
3. **Informações Necessárias:**
   - Último lançamento no prefixo RHI (painel [07.01]).
   - Data de pagamento.
   - Natureza do pagamento: `007001001` (salário) ou `007001002` (adiantamento).
   - Histórico: `Folha (Salário/Adiantamento) (mm/aa)`.
4. **Baixar e Executar:**
   - Baixe o repositório como ZIP.
   - Execute o programa na pasta `dist` como administrador.
   - Selecione o arquivo Excel e salve o novo arquivo com "- Formatado".
5. **Preenchimento das Informações:**
   - Forneça o número maior que o último lançamento (E2_NUM).
   - Informe a natureza, data de pagamento e histórico formatado.

## Validações

- Certifique-se de que o Excel tenha apenas uma aba.
- Valide o formato da folha de pagamento.
- Use o programa "**PROCV Fornecedor**" para preencher os códigos de fornecedor.

## Importação

- Siga as instruções detalhadas [aqui](https://docs.google.com/document/d/15Df5Kwursnd_0qNLkxGTVh5r1F04CFqWrY0_7P-r5A4/edit?usp=sharing).

## Estrutura do Repositório

```plaintext
Bred-Capas/
├── dist/
│   ├── Folha de pagamento/
│   └── Procv Fornecedor/
├── README.md
└── ...
```

## Suporte

Para dúvidas, entre em contato com o suporte da empresa ou abra uma issue no repositório.
