# 🚚 VBA – Automação de Criação de Ordens de Transporte

Macros desenvolvidas em **VBA (Visual Basic for Applications)** para automação de processos no **Microsoft Excel**, integradas ao **SAP GUI**.  
O projeto permite criar ordens de transporte, gerenciar remessas e atualizar planilhas automaticamente, reduzindo erros e aumentando a produtividade.

---


## 🚀 Como Utilizar

### 🔹 Usar a versão pronta
1. Acesse a pasta [`build/`](./build).  
2. Baixe o arquivo `.xlsm`.  
3. Abra no Excel e habilite as macros.  

### 🔹 Usar apenas o código-fonte
1. Abra o Excel e pressione `ALT + F11` para acessar o Editor VBA.  
2. Vá em **Arquivo > Importar arquivo**.  
3. Selecione os módulos em [`src/`](./src).  
4. O código será importado automaticamente para o seu projeto VBA.  

---

## 🛠️ Desenvolvimento

- **Automação SAP GUI** → Criação de transportes diretamente no SAP.  
- **Validação de dados** → Garantia de consistência de notas fiscais, materiais e parceiros.
- **Compatibilidade** → Testado no Excel 2016, 2019 e Microsoft 365.  

### 🔧 Contribuir
1. Faça um **fork** deste repositório.  
2. Clone para sua máquina:  
   ```bash
   git clone https://github.com/usuario/TransportAutomation.git