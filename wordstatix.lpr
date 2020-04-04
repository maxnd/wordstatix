// ***********************************************************************
// ***********************************************************************
// WordStatix 2.x
// Author and copyright: Massimo Nardello, Modena (Italy) 2016 - 2020.
// Free software released under GPL licence version 3 or later.

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version. You can read the version 3
// of the Licence in http://www.gnu.org/licenses/gpl-3.0.txt
// or in the file Licence.txt included in the files of the
// source code of this software.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// ***********************************************************************
// ***********************************************************************
program wordstatix;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, tachartlazaruspkg, Unit1, Unit2, Copyright
  {$ifdef Linux}, clocale{$endif}
  {$ifdef Darwin}, clocale, iosxlocale{$endif}
  { you can add units after this };

{$R *.res}

begin
  Application.Scaled:=True;
  Application.Title:='WordStatix';
  RequireDerivedFormResource:=True;
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TfmCopyright, fmCopyright);
  Application.CreateForm(TfmInput, fmInput);
  Application.Run;
end.

