// How to Use
// ScCheque.exe "001" "150,00" "Curitiba" "3112018"
//ou
// ScCheque.exe "001" "150,00" "Curitiba" "3112018" "Jose Silva"
// ou
// ScCheque.exe "001" "150,00" "Curitiba" "3112018" "Jose Silva" "Bom p/ 01/01/2019"
// ou
// ScCheque.exe

unit sc_cheque;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  portas : array[0..8] of string = ('COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'LPT1', 'LPT2', 'LPT3');
  Form1: TForm1;
  i, retornoPorta, retornoImpressao : integer;
  msg_validacao, msg, data, cidade, favorecido, porta, banco, valor: string;
  helpCommand : string = 'Parâmetros Inválidos. Tente o comando: ScCheque.exe "001" "150,00" "Curitiba" "3112018" "Jose Silva" "Bom p/ 01/01/2019"';
implementation
  // BEMATECH funcitions
  function Bematech_DP_IniciaPorta(Porta: string): integer; stdcall; far; external 'BemaDP32.dll';
  function Bematech_DP_FechaPorta: integer; stdcall; far; external 'BemaDP32.dll';
  function Bematech_DP_ImprimeCheque(Banco: string; Valor: string; Favorecido: string; Cidade: string; Data: string; Mensagem: string): integer; stdcall; far; external 'BemaDP32.dll';

  procedure exitProgram();
    begin
      Halt(4);
    end;

  procedure alert(const msg: String);
    begin
      MessageDlg(msg, mtInformation, [mbOk], 0);
    end;
    
  function IfNull( const Value, Default : OleVariant ) : OleVariant;
  begin
    if (Value = NULL) or (Value = '') then
      Result := Default
    else
      Result := Value;
  end;

  procedure setParams();
    begin
      // PARAMS
      banco      := paramstr(1); // params que chegam quando usuário executa o .exe por linha de comando
      valor      := paramstr(2);
      cidade     := paramstr(3);
      data       := paramstr(4);
      favorecido := paramstr(5);
      msg        := paramstr(6);

      // impressao teste quando não enviamos parametros
      if banco = '' then
        begin
          banco      := IfNull(banco, '341'); // banco do brasil
          valor      := IfNull(valor, '150,00');
          favorecido := IfNull(favorecido, 'Jose da Silva');
          cidade     := IfNull(cidade, 'Curitiba');
          data       := IfNull(data, '15102003');
          msg        := IfNull(msg, '');
        end;

      if (banco = '') or (valor = '') or (cidade = '') or (data = '') then
        begin
          alert(helpCommand);
          exitProgram();
        end;
    end;

  procedure abrirPorta();
    begin
      for i := 0 to (length(portas)-2) do
        begin
        porta := portas[i];
        //alert(porta);
        retornoPorta := Bematech_DP_IniciaPorta(porta);
        if retornoPorta = 1 then
          Break;
      end;
    end;

  procedure imprimir(Banco: string; Valor: string; Favorecido: string; Cidade: string; Data: string; Mensagem: string);
    begin
      retornoImpressao := Bematech_DP_ImprimeCheque( pchar( banco ), pchar( valor ), pchar( favorecido ), pchar( cidade ), pchar( data ), pchar( msg ) );
      if retornoImpressao = 0 then
        alert('Erro de Comunicação com a Impressora. Verifique.');
      if retornoImpressao = 1 then
        //alert(Imprimindo...');
      if retornoImpressao = -2 then
        alert(helpCommand);
      if retornoImpressao = -3 then
          alert('Banco '+banco+' não encontrado no arquivo BEMAPDP32.ini.');
    end;

  procedure tentarImprimir();
    begin
      abrirPorta();
      if retornoPorta = 0 then
        alert('Erro ao abrir portas');
      if retornoPorta = 1 then // sucesso
        imprimir(banco, valor, favorecido, cidade, data, msg);
      Bematech_DP_FechaPorta();
    end;

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin

  setParams();
  tentarImprimir();
  exitProgram();

end;

end.
