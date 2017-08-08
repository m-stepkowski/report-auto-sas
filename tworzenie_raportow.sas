options sastrace=',,,d' sastraceloc=saslog;

%let projekty = D:\projekt\raporty\;

libname wyj_exc "&projekty.excel" compress=yes;
libname wyj_html "&projekty.html" compress=yes;
libname wyj_pdf "&projekty.pdf" compress=yes;

/* Przygotowanie polaczonej zbioru transactions i production - dokonanie kategoryzacji */

data _null_;
	set wyj.podzialy_int_niem_css end = eof;
	if _n_ = 1 then do;
		call execute("data po_kate;");
		call execute("set wej.production;");
		call execute("if product='css' then do;");
	end;
		call execute("if "||strip(war)||" then "||strip(zmienna)||" = "||strip(put(grp,2.))||";"); 
	if eof = 1 then call execute("end; run;");
run; 

data _null_;
	set wyj.podzialy_int_niem_ins end = eof;
	if _n_ = 1 then do;
		call execute("data po_kate;");
		call execute("set po_kate;");
		call execute("if product='ins' then do;");
	end;
		call execute("if "||strip(war)||" then "||strip(zmienna)||" = "||strip(put(grp,2.))||";"); 
	if eof = 1 then call execute("end; run;");
run;

/*proc sql noprint;*/
/*	create table analiza_wejscie as*/
/*	select a.*, b.* from po_kate a*/
/*	left join wej.transactions b on b.aid=a.aid;*/
/*quit;*/


/* Utworzenie wyniku o danej nazwie dla danych warunkow ze zbioru wejsciowego vin - zbior SASowy */

%macro raport_dane(nazwa_zbioru);

proc means data=vin noprint nway;
	class fin_period seniority;
	var vin3;
	output out=vintagr(drop=_freq_ _type_) n()=production mean()=vintage3;
	format vintage3 nlpct12.2;
run;

proc means data=vin noprint nway;
	class fin_period;
	var vin3;
	output out=production(drop=_freq_ _type_) n()=production;
	where seniority=0;
run;

proc transpose data=vintagr out=vintage prefix=months_after_;
	by fin_period;
	var vintage3;
	id seniority;
run;

data prediction;
	format data_spr yymmdd10.;
	set vintage (keep=fin_period months_after_12);
	data_spr = input(fin_period, yymmn6.);
	drop fin_period;
run;	
	
proc arima data=prediction;
  identify var=months_after_12;
  estimate p=1 ;
  forecast lead=12 interval=month id=data_spr out=prediction1;
quit;

data prediction1 (rename=(months_after_12=forecast_12));
	format fin_period $10.;
	set prediction1;
	if data_spr >= "01DEC2007"d then months_after_12 = forecast;
	else months_after_12 = .;
	fin_period = put(data_spr, yymmn6.);
	keep fin_period months_after_12;
run;
	
data wyj_rap.&nazwa_zbioru;
	retain fin_period production forecast_12 months_after_0-months_after_35;
	merge vintage (drop=_name_) production prediction1;
	by fin_period;
run;

%mend raport_dane;


/* Utworzenie raportu EXCELowego na podstawie danych SASowych */

%macro raport_excel(nazwa_zbioru, ile_due, nazwa_pdf);

%let wej_exc = D:\projekt\dane;
%let wyj_exc = D:\projekt\raporty\excel;

/* Otwarcie Excela */

filename ddedata dde 'excel|system';
data _null_;
	length fid rc start stop time 8;
	fid=fopen('ddedata','s');
	if (fid le 0) then do;
		rc=system('start excel');
		start=datetime();
		stop=start+10;
		do while (fid le 0);
			fid=fopen('ddedata','s');
			time=datetime();
			if (time ge stop) then do; 
				fid=1;
			end;
		end;
	end;
	rc=fclose(fid);
run;

/* Otwarcie wzoru arkusza */

data _null_;
  length plik $300;
  filename ddedata dde 'excel|system'; 
  file ddedata;
  plik='[open("'||"&wej_exc.\"||'temp.xlsx")]';
  put plik;
run;

/* Insert danych */

data _null_;  
	filename ddedata dde 'excel|data!w2k1:w37k39';
	file ddedata; 
	set wyj_rap.&nazwa_zbioru;
	put fin_period production forecast_12 months_after_0-months_after_35;
run; 
 

/* Zapis arkusza */

data _null_;
	length plik $300;
	filename ddedata dde 'excel|system'; 
	file ddedata; 
	plik='[save.as("'||"&wyj_exc.\"||"&nazwa_zbioru."||'", 50)]';
	put plik;
	put '[close()]';
run;

/* Zamkniecie pliku */

data _null_;
	filename ddedata dde 'excel|system'; 
	file ddedata; 
	put '[quit()]';
run;

%mend raport_excel;


/* Utworzenie raportu HTMLowego na podstawie danych SASowych */

%macro raport_html(nazwa_zbioru);

%let projekty_html = D:\projekt\raporty\html\;


ods listing close;

   ods html path="&projekty_html." gpath="&projekty_html.\png" (url="png/") file="&nazwa_zbioru..html" Style=HTMLBlue;
	title;
	ods html text="<center><h1>Raport Vintage dla (&nazwa_pdf.)</h1></center><br>";

	ods html text='<center><h2>Linki:</h2></center><br>';

	ods html text="<table><tr>";
	ods html text="<td><p align=""center""><a href='../excel/&nazwa_zbioru..xlsb'>Pobierz Raport Excel</p><td>
					<td><p align=""center""><a href='../pdf/&nazwa_zbioru._rap1.pdf'>Pobierz Raport PDF</p></td>" ;
	ods html text="</tr></table>";

	ods html text='<br><br>';

	ods html text='<center><h2>Wykres:</h2></center><br>';

	ods graphics on / imagename="&nazwa_zbioru." ;  
        proc sgplot data=wyj_rap.&nazwa_zbioru. (where=(mod(month(input(fin_period, yymmn6.)), 3) = 0));
			vbar fin_period / response=production y2axis;
   			vline fin_period / response=months_after_3;
			vline fin_period / response=months_after_6;
			vline fin_period / response=months_after_9;
			vline fin_period / response=months_after_12;
			vline fin_period / response=forecast_12;
			yaxis min=0 label="Vintage &j.+";
        run;
   	ods graphics off;

	ods html text='<br><br>';

	ods html text='<center><h2>Dane:</h2></center><br>';

	proc print data=wyj_rap.&nazwa_zbioru;
	var fin_period production forecast_12 months_after_3 months_after_6 months_after_9 months_after_12;
 	run;


   ods html close;
ods listing;

%mend;


/* Utworzenie raportu PDF na podstawie danych SASowych */

%macro raport_pdf(nazwa_zbioru, nazwa_pdf);

%let projekty_pdf = D:\projekt\raporty\pdf\;


ods listing gpath="&projekty_pdf.";
	goptions device=pdf colors=(black) rotate=landscape;
	ods html close;
	ods graphics on / imagefmt=pdf imagename="&nazwa_zbioru._rap" ;

		PROC TEMPLATE;
		DEFINE STATGRAPH mixed;
			BEGINGRAPH;
				ENTRYTITLE "Raport Vintage dla (&nazwa_pdf)";
				LAYOUT LATTICE / ROWS= 2 ROWGUTTER=50;
					LAYOUT OVERLAY / YAXISOPTS = (LABEL= 'Production') XAXISOPTS = (LABEL= 'Data');
						BARCHART Y = production X = fin_period;
					ENDLAYOUT;
					LAYOUT OVERLAY / YAXISOPTS = (LABEL= 'Vintage') XAXISOPTS = (LABEL= 'Data') ;
						LINECHART category=fin_period response=months_after_3 /LINEATTRS=(color=blue) NAME="months_after_3";
						LINECHART category=fin_period response=months_after_6 /LINEATTRS=(color=green) NAME="months_after_6";
						LINECHART category=fin_period response=months_after_9 /LINEATTRS=(color=purple) NAME="months_after_9";
						LINECHART category=fin_period response=months_after_12 /LINEATTRS=(color=black) NAME="months_after_12";
						LINECHART category=fin_period response=forecast_12 /LINEATTRS=(color=red) NAME="forecast_12";
						DISCRETELEGEND "months_after_3" "months_after_6" "months_after_9" "months_after_12" "forecast_12";
					ENDLAYOUT;
				ENDLAYOUT;
			ENDGRAPH;
		END;
		RUN;

		PROC SGRENDER DATA = wyj_rap.&nazwa_zbioru. (where=(mod(month(input(fin_period, yymmn6.)), 3) = 0)) TEMPLATE = mixed;
		RUN;
		
	ods graphics off;
ods listing ;

%mend;


/* Makro tworzace przygotowujace dane i wszystkie raporty */

%macro raportowy_potwor(nazwa_zbioru, nazwa_pdf, ile_due);

	%raport_dane(&nazwa_zbioru);

	%raport_excel(&nazwa_zbioru, &ile_due, &nazwa_pdf);

	%raport_html(&nazwa_zbioru);

	%raport_pdf(&nazwa_zbioru, &nazwa_pdf);

%mend raportowy_potwor;


/* ---------------- Przygotowanie kategorii do tworzenia raportow ------------------ */


/* Liczba opoznien poczawszy od 1. Jezeli liczba_due 3, to obliczany jest dla due=1,2,3 */
%let liczba_due = 3;


/* Produkty, ktore sa w zbiorze plus TOTAL */
data produkty_do_raportow;
	format product $10. product_warunek $100.;
	if product = "" then delete;
run;

proc sql noprint;
	insert into produkty_do_raportow 
	select distinct product, strip("where product='"||strip(product)||"'") from wej.Transactions;

	insert into produkty_do_raportow values ("total", "where 1=1");
quit;

data produkty_do_raportow;
	set produkty_do_raportow;
	lp = _N_;
run;



/* ---------------- Czesc tworzaca raporty --------------- */

%macro tworz_raporty;

data due_do_raportow;
	format due 1. due_warunek $100.;
	if due = "" then delete;
run;

%do i=1 %to &liczba_due;
	proc sql noprint;
		insert into due_do_raportow values (&i, "due_installments>=&i" );
	quit;
%end;

data due_do_raportow;
	set due_do_raportow;
	lp = _N_;
run;


/* Pobranie informacji o liczbie produktow, due i zmiennych */

proc sql noprint;
	select count(*) into: _ile_prod from produkty_do_raportow;
	select count(*) into: _ile_due from due_do_raportow;
quit;

%do i=1 %to &_ile_prod;

	data _null_;
		set produkty_do_raportow (where=(lp=&i));
		Call Symput("product", strip(product));
		Call Symput("product_warunek", strip(product_warunek));
	run;

	%if "&product" = "css" %then %do;
		proc sql noprint;
			create table zmienne_do_raportow as select name, grupa, monotonic() as lp from wyj.vin_out_zmienne_css
			where name ne "";
			create table podzialy_int_niem as select * from wyj.podzialy_int_niem_css;
		quit;
	%end;
	%else %do;
		proc sql noprint;
			create table zmienne_do_raportow as select name, grupa, monotonic() as lp from wyj.vin_out_zmienne_ins
			where name ne "";
			create table podzialy_int_niem as select * from wyj.podzialy_int_niem_ins;
		quit;
	%end;

	proc sql noprint;
		select count(*) into: _ile_vars from zmienne_do_raportow;	
	quit;

	%if "&product" = "total" %then %do;
		proc sql noprint;
			create table analiza_wejscie as
			select * from wej.Transactions
			where aid in
				(
					select aid from wej.Production
				)
		;quit;
	%end;
	%else %do;
		proc sql noprint;
			create table analiza_wejscie as
			select * from wej.Transactions
			where aid in
				(
					select aid from wej.Production
					where product = "&product"
				)
		;quit;
	%end;

	/* 1. Raporty dla produktu, due */

	%do j=1 %to &_ile_due;

		data _null_;
			set due_do_raportow (where=(lp=&j));
			Call Symput("due_warunek", strip(due_warunek));
		run;

		%let nazwa_pdf = product_&product. due&j.;
		%let nazwa_kon = &product._due&j.;
		%put &nazwa_pdf;

		data vin;
			set analiza_wejscie;
			seniority=intck('month',input(fin_period,yymmn6.),input(period,yymmn6.));
			vin3=(&due_warunek);
			output;
			if status in ('B','C') and period<='200812' then do;
				n_steps=intck('month',input(period,yymmn6.),input('200812',yymmn6.));
				do i=1 to n_steps;
					period=put(intnx('month',input(period,yymmn6.),1,'end'),yymmn6.);
					seniority=intck('month',input(fin_period,yymmn6.),input(period,yymmn6.));
					output;
				end;
			end;
			keep fin_period vin3 seniority;
		run;

		%raportowy_potwor(&nazwa_kon, &nazwa_pdf, &j);

		/* 3. Raporty dla produktu, due, zmiennej */

		%do k=1 %to &_ile_vars;

			data _null_;
				set zmienne_do_raportow (where=(lp=&k));
				Call Symput("zmienna", strip(name));
				Call Symput("var_grupa", strip(grupa));
				Call Symput("var_name_warunek", strip(name));
			run;

			%if "&product" = "total" %then %do;
				proc sql noprint;
					create table analiza_wejscie1 as
					select * from wej.Transactions
					where aid in
						(
							select aid from wej.Production
							where not missing(&zmienna)
						)
				;quit;
			%end;
			%else %do;
				proc sql noprint;
					create table analiza_wejscie1 as
					select * from wej.Transactions
					where aid in
						(
							select aid from wej.Production
							where product = "&product" and not missing(&zmienna)
						)
				;quit;
			%end;


			%let nazwa_pdf = product_&product. due_&j. zmienna_&zmienna.;
			%let nazwa_kon = &product._due&j._&var_grupa.&k.;

			data vin;
				set analiza_wejscie1;
				seniority=intck('month',input(fin_period,yymmn6.),input(period,yymmn6.));
				vin3=(&due_warunek);
				output;
				if status in ('B','C') and period<='200812' then do;
					n_steps=intck('month',input(period,yymmn6.),input('200812',yymmn6.));
					do i=1 to n_steps;
						period=put(intnx('month',input(period,yymmn6.),1,'end'),yymmn6.);
						seniority=intck('month',input(fin_period,yymmn6.),input(period,yymmn6.));
						output;
					end;
				end;
				keep fin_period vin3 seniority;
			run;

			%raportowy_potwor(&nazwa_kon, &nazwa_pdf, &j);

			proc sql noprint;
				create table zmienna_kat as select war, zmienna, grp, monotonic() as lp from podzialy_int_niem
				where lowcase(zmienna) = strip(lowcase("&var_name_warunek"));

				select count(*) into: liczba_kat from zmienna_kat;
			quit;

			/* 4. Raporty dla produktu, due, zmiennej i kategorii */

			%do l=1 %to &liczba_kat;

				data _null_;
					set zmienna_kat (where=(lp=&l));
					Call Symput("zmienna", strip(zmienna));
					Call Symput("var_nr_kat", strip(grp));
					Call Symput("zmienna_warunek", strip(war));
				run;

				%if "&product" = "total" %then %do;
					proc sql noprint;
						create table analiza_wejscie2 as
						select * from wej.Transactions
						where aid in
							(
								select aid from wej.Production
								where &zmienna_warunek.
							)
					;quit;
				%end;
				%else %do;
					proc sql noprint;
						create table analiza_wejscie2 as
						select * from wej.Transactions
						where aid in
							(
								select aid from wej.Production
								where product = "&product" and &zmienna_warunek.
							)
					;quit;
				%end;
				
				%let nazwa_pdf = product_&product. due_&j. zmienna_&zmienna. kat_&var_nr_kat.;
				%let nazwa_kon = &product._due&j._&var_grupa.&k._kat&var_nr_kat.;

				data vin;
					set analiza_wejscie2;
					seniority=intck('month',input(fin_period,yymmn6.),input(period,yymmn6.));
					vin3=(&due_warunek);
					output;
					if status in ('B','C') and period<='200812' then do;
						n_steps=intck('month',input(period,yymmn6.),input('200812',yymmn6.));
						do i=1 to n_steps;
							period=put(intnx('month',input(period,yymmn6.),1,'end'),yymmn6.);
							seniority=intck('month',input(fin_period,yymmn6.),input(period,yymmn6.));
							output;
						end;
					end;
					keep fin_period vin3 seniority;
				run;

				%raportowy_potwor(&nazwa_kon, &nazwa_pdf, &j);

			%end;

		%end;

	%end;

%end;

%mend tworz_raporty;


%tworz_raporty;