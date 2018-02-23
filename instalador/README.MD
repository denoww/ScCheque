# Instalação no Windows

  1 - Copie todos os arquivos que estão dentro da pasta DLL para C:\Windows (caso já tenha a impressora funcionando com outro software você DEVE PULAR este passo

  2 - Copie ScCheque.exe para a pasta C:
  
  3 - Execute o arquivo instalar_navegador.reg 
      
    se este ultimo passo der erro, nos notifique para gerarmos o mesmo arquivo a partir do seu computador

  Só funcinoa no windows (restrição da empresa Bematech)

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

# Como adicionar cheques de novos Bancos

    Para adicionar novos cheques de bancos, ou modificar existentes será necessário editar o arquivo C:\Windows\BEMADP32.ini
    Não é necessário reiniciar o computador
    Caso as modificações não estejam funcionando:
      APAGUE o arquivo C:\Users\NOME DO SEU USUário\AppData\Local\VirtualStore\Windows\BEMADP32.ini
      Ou: use a busca do windows e apague os arquivos BEMADP32.ini que NÃO estão dentro de C:\Windows

