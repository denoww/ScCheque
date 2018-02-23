# Instalação

  Copie todos os arquivos que estão dentro da pasta DLL para C:\Windows

  Copie ScCheque.exe para a pasta C:
  
  Execute o arquivo instalar_navegador.reg 
    (se esse passo der erro nos notifique para gerarmos o mesmo arquivo a partir do seu computador

# Como Imprimir

## Imprimir pelo navegador

  Código Html
  ```html
<a href='ScCheque:"001" "150,00" "Curitiba" "31122018" "Flavio Barros" "Bom p/ 01/01"'>Imprimir</a>
<a href='ScCheque:"001" "150,00" "Curitiba" "31122018" " " " "'>Imprimir Campos Básicos</a>
<a href='ScCheque:"001" "150,00" "Curitiba" "31122018" " " "Bom p/ 01/01"'>Sem Favorecido</a>
```

  


## Imprimir no Windows

### Impressão Básica

   Pressione: tecla windows + R

   C:\ScCheque.exe "001" "150,00" "Curitiba" "3112018" "Jose Silva" " " " "

   Pression ENTER

### Cheque Nominal e Mensagem

  Pressione: tecla windows + R 
  
  Cole o código abaixo
  
  ```html
C:\ScCheque.exe "001" "150,00" "Curitiba" "3112018" "Jose Silva" "Bom p/ 01/01/2019"
```

  Pression ENTER

