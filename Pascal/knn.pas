Program Trabalho_KNN;

uses crt;

type PontElemento=^Elemento;

     Elemento= record
               Resultado:real;
               Distancia:real;
               Prox:PontElemento;
               end;

     PontDescritor=^Descritor;

     Descritor= record
                inicio,fim:PontElemento;
                n:integer;
                End;

     matriz=array[1..900,1..9] of real;


var
    vs:array [1..50] of string;

    Mat,M1:matriz;

    i,j,n,lin:integer;

    Desc:PontDescritor;

    Dist:real;

    Atual:PontElemento;

Procedure Armazena(var s:string;lin:integer);

var st:string;
    l,code:integer;

Begin

j:=1;

s:=s+ ' ';

l:=length(s);

st:='';

For i:=1 to l do

    Begin

    if s[i] <> ' ' then

    st:=st+ s[i]

    Else

        Begin

        vs[j]:=st;

        inc(j);

        st:='';

        End;

    End;

For j:=1 to 9 do
    Begin
    val(vs[j],mat[lin,j],code);
    End;
End;



Procedure LeArquivo;

var Dados:Text;
    s:string;

Begin

Assign(Dados,'C:/Diabetes/Pima1.txt');

reset(Dados);

lin:=1;

While not eof(Dados) do

      Begin

      readln(Dados,s); {Le todos os dados de cada linha}

      Armazena(s,lin);

      Inc(lin);
      End;

Close(Dados);


End;

Procedure Ler_dados;

Begin
For j:=1 to 8 do

    Readln(M1[1,j]);

End;

Procedure Calcular_Distancia(lin:integer;var dist:real);

var quadrado:real;

    j:integer;

Begin

Dist:=0;

For j:=1 to  8 do
    Begin

    quadrado:=0;

    quadrado:=sqr(M1[1,j]-Mat[lin,j]); {Tomar cuidado com o lin, pois esta somando soh M1}

    dist:=quadrado+dist;

    End;

dist:=sqrt(dist);


End;

Procedure Criar_Lista;

Begin

New(Desc);

Desc^.inicio:=nil;

Desc^.n:=0;

Desc^.fim:=nil;


End;

Procedure insere(Desc:PontDescritor;dist:real;resultado:real);

var p,atual:PontElemento;

Begin


New(p);

p^.distancia:=dist;

p^.resultado:=resultado;


If  Desc^.n=0 then

    Begin

    Desc^.inicio:=p;

    Desc^.fim:=p;

    Desc^.n:=1;

    p^.prox:=nil;

    End

Else
    Begin
    p^.prox:=Desc^.inicio;
    Desc^.inicio:=p;
    Desc^.n:=Desc^.n+1;
    End;



End;

Procedure Ordena(n:integer;var Desc:PontDescritor);

var i,j:integer;
    atual,anterior,aux:PontElemento;
Begin

Atual:=Desc^.inicio;
Anterior:=Atual;
Atual:=Atual^.Prox;

For i:=n downto 2 do
    Begin
    For j:=1 to i-1 do
        Begin
        if Anterior^.distancia>Atual^.distancia then
            Begin
            New(Aux);
            Aux^.distancia:=Anterior^.distancia;
            Aux^.resultado:=Anterior^.resultado;
            Anterior^.distancia:=Atual^.distancia;
            Anterior^.resultado:=Atual^.resultado;
            Atual^.distancia:=Aux^.distancia;
            Atual^.resultado:=Aux^.resultado;
            Dispose(Aux);
            End;
        Atual:=Atual^.Prox;
        Anterior:=Anterior^.Prox;
        End;
    Atual:=Desc^.inicio;
    Anterior:=Atual;
    Atual:=Atual^.Prox;
    End;

Atual:=Desc^.inicio;

For i:=1 to 3 do
    Begin
    Writeln(Atual^.distancia:3:3,'  ',Atual^.resultado:1:0);
    Atual:=Atual^.Prox;
    End;




End;

Begin

clrscr;

LeArquivo;

Criar_Lista;

Ler_Dados;

For i:=1 to 200 do

    Begin

    Calcular_Distancia(i,Dist);

    Insere(Desc,Dist,Mat[i,9]);

    End;


Atual:=Desc^.inicio;

n:=200;

Ordena(n,desc);


readkey;

End.
