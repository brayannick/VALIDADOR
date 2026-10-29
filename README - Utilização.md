# Validador de Alunos

Confere se um novo cadastro pode ser duplicado, usando uma planilha (XLSX ou CSV)
exportada do sistema da instituição. O programa não acessa o sistema original e
não grava nada na planilha.

## Instalação (Windows)

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python main.py
```

No Linux, pode ser necessário instalar o pacote `python3-tk`.

## Planilha

- Colunas obrigatórias: `nome` e `telefone` (maiúsculas não importam).
- A coluna "Linha" nos resultados indica a linha da planilha onde o cadastro está.
- Um exemplo fictício está em `dados/alunos_exemplo.csv`. Não use dados reais em demonstrações.

## Como usar

1. Clique em **Selecionar planilha** (a última base usada é reaberta automaticamente).
2. Informe nome e telefone e clique em **Verificar duplicidade** (ou tecle Enter).
3. Confira os resultados no sistema original antes de cadastrar.
4. **Atualizar base** relê o arquivo. Se a planilha for substituída ou salva de novo,
   o programa também a recarrega sozinho na próxima verificação.

## Regras de comparação

| Classificação | Quando aparece |
|---|---|
| Forte | Nome e telefone coincidem; ou telefone coincide e o nome é semelhante |
| Conferir | Só o telefone coincide (pode ser número compartilhado); ou só o nome coincide |
| Possível | Nome semelhante (acima do limite de similaridade) ou nome contido em outro cadastro (ex.: "Maria Silva" em "Maria Oliveira Silva") |

## Padronização de telefones

- Remove símbolos, o prefixo `00`, o código do país `55` e o zero antes do DDD.
- Telefone sem DDD recebe o **DDD padrão** (campo na tela, inicial 88).
- A comparação usa os 8 últimos dígitos, então o número antigo de 8 dígitos
  coincide com a versão de 9 dígitos. Se os dois tiverem DDD, ele precisa ser igual.

## Ajustes

- **Similaridade mínima de nomes**: inicial 86%. Ajuste após testar com dados representativos.
- DDD padrão e similaridade ficam salvos em `config.json`, ao lado do programa.
- `validador.log` registra apenas eventos técnicos (contagens e erros), nunca nomes ou telefones.
- A tela avisa quando a planilha tem 7 dias ou mais (`DIAS_ALERTA_BASE` em `main.py`).

## Segurança

A planilha contém dados pessoais: guarde-a em local institucional com acesso restrito,
não a envie por e-mail pessoal ou serviços não autorizados e não inclua dados reais em
capturas de tela ou relatórios de teste.

## Gerar executável (opcional)

```bash
python -m pip install pyinstaller
python -m PyInstaller --onefile --windowed main.py
```

O executável fica em `dist`. Teste-o no computador de destino antes de distribuir.

## Limitações

- Não impede que alguém cadastre direto no sistema original sem passar pela validação.
- Se a planilha estiver desatualizada, cadastros recentes não serão detectados.
- Resultados são sugestões para conferência, não confirmações de identidade.
