<html><head></head>
  <body>
  <p align="center"><b>PLANILHAR NOTAS EM PDF</b></a>
  <p>O projeto consistiu em planilhar dados de todas as notas de empenho em PDF de 2020 e 2021. Para isso, utilizou-se o VBA e a biblioteca Pyautogui do Python. </p>
  <p>Como se tratava de uma vasta documentação em PDF, a utilização de um bot de movimentação do mouse facilitou as atividades repetitivas de conversão dos PDFs. Optou-se por fazer download de todos os PDFs e convertê-los para Excel via site EasyPDF e o bot fez as movimentações no mouse (vide gif abaixo) </p>
   
<p align="center"><img src="https://im3.ezgif.com/tmp/ezgif-3-9db99eb2a1.gif">
  
<br>
  <p>Posteriormente, o conteúdo do Excel é copiado e colado na planilha "Separar em abas" que separa, via macro, cada empenho em uma aba diferente.</p>
  <img src="https://github.com/RenataVerasVenturim/relacionarempenhos2021/assets/129551549/334e9085-7171-496e-b57c-bdcc32c4b06e">

  <br>
  <p>Na planilha "Planilha teste unificar", as planilhas separadas em abas são importadas e as macros são executadas de extração dos dados necessários de cada empenho (50 por vez , visto que o VBA não suportaria mais que isso, após testes realizados. Times foram incluídos)</p>
  <img src="https://github.com/RenataVerasVenturim/relacionarempenhos2021/assets/129551549/31cadae9-b615-4575-8bf0-ef98127bbebe">
  <p>Após isso, em "EMPENHOS2021-Consolidado" é acionado o botão de "Atualizar" que aciona macros que unificam as informações de todas as abas, conforme abaixo:</p>
<br>
  <p><img src="https://im3.ezgif.com/tmp/ezgif-3-ed43b51d8e.gif">
    
  </body>
</html>
