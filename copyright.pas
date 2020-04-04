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
unit Copyright;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, LCLIntf;

type

  { TfmCopyright }

  TfmCopyright = class(TForm)
    bnCopyright: TButton;
    imImage: TImage;
    lbCopyrightSubTitle: TLabel;
    lbCopyrightAuthor1: TLabel;
    lbCopyrightAuthor2: TLabel;
    lbCopyrightName: TLabel;
    lbSite: TLabel;
    meCopyrightText: TMemo;
    tmAlarmForm: TTimer;
    procedure bnCopyrightClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure lbSiteClick(Sender: TObject);
    procedure tmAlarmFormTimer(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmCopyright: TfmCopyright;

implementation

{$R *.lfm}

{ TfmCopyright }

procedure TfmCopyright.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  // Stop the timer anyway
  tmAlarmForm.Enabled := False;
end;

procedure TfmCopyright.FormCreate(Sender: TObject);
begin
  // Set color
  {$ifdef Linux}
  fmCopyright.Color := clWhite;
  {$endif}
  {$ifdef Windows}
  fmCopyright.Color := clWhite;
  {$endif}
end;

procedure TfmCopyright.bnCopyrightClick(Sender: TObject);
begin
  // Quit
  Close;
end;

procedure TfmCopyright.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Quit
  if key = 27 then
    Close;
end;

procedure TfmCopyright.FormShow(Sender: TObject);
begin
  // Start fading
  fmCopyright.AlphaBlendValue := 0;
  tmAlarmForm.Enabled := True;
end;

procedure TfmCopyright.lbSiteClick(Sender: TObject);
begin
  // Open the website
  OpenURL('http://sites.google.com/site/wordstatix/');
end;

procedure TfmCopyright.tmAlarmFormTimer(Sender: TObject);
begin
  // Show the form fading
  if fmCopyright.AlphaBlendValue < 255 then
  begin
    fmCopyright.AlphaBlendValue := fmCopyright.AlphaBlendValue + 1;
  end
  else
  begin
    tmAlarmForm.Enabled := False;
  end;
end;

end.
