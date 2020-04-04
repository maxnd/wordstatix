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
unit unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Grids,
  StdCtrls, ExtCtrls, ComCtrls, Menus, LazUTF8, TAGraph, TASources, TASeries,
  TATools, IniFiles, Clipbrd, zipper, LCLIntf, CheckLst, Types, LazFileUtils,
  gettext, DefaultTranslator;

type

  TRecWordList = record
    stRecWord: string;
    iRecFreq: integer;
    stRecWordPos: string;
    stRecBookPos: string;
  end;

  TRecWordTextual = record
    stRecWordTextual: string;
  end;

  { TfmMain }

  TfmMain = class(TForm)
    apAppProp: TApplicationProperties;
    bnDeselConc: TButton;
    bnFindFirst: TButton;
    bnFindNext: TButton;
    bnReplace: TButton;
    bnReplaceAll: TButton;
    bnDeselDiag: TButton;
    cbCollatedSort: TCheckBox;
    cbComboDiag1: TComboBox;
    cbComboDiag2: TComboBox;
    cbComboDiag3: TComboBox;
    cbComboDiag4: TComboBox;
    cbComboDiag5: TComboBox;
    cbSkipNumbers: TCheckBox;
    chChartBarSeries4: TBarSeries;
    chChartBarSeries5: TBarSeries;
    csChartSource4: TListChartSource;
    csChartSource5: TListChartSource;
    chChart: TChart;
    chChartBarSeries1: TBarSeries;
    chChartBarSeries2: TBarSeries;
    chChartBarSeries3: TBarSeries;
    clDiagBookmark: TCheckListBox;
    csChartSource2: TListChartSource;
    csChartSource3: TListChartSource;
    edFltStart: TEdit;
    edFltEnd: TEdit;
    edFind: TEdit;
    edLocate: TEdit;
    edReplace: TEdit;
    edWordsCont: TEdit;
    lbDiagBookmark: TLabel;
    lbDiag3: TLabel;
    lbDiag2: TLabel;
    lbDiag1: TLabel;
    lbBookmarks: TListBox;
    lbDiag4: TLabel;
    lbDiag5: TLabel;
    lbLocate: TLabel;
    lbReplace: TLabel;
    lbFind: TLabel;
    lbFltStart: TLabel;
    lbFltEnd: TLabel;
    lbWordsCont: TLabel;
    lbContext: TLabel;
    lbNoWord: TLabel;
    lsContext: TListBox;
    csChartSource1: TListChartSource;
    miFileCopyright: TMenuItem;
    miFileManual: TMenuItem;
    miLineManual: TMenuItem;
    miLine1: TMenuItem;
    miFileSaveConc: TMenuItem;
    miFileOpenConc: TMenuItem;
    miStatisticSortFreq: TMenuItem;
    miStatisticSortName: TMenuItem;
    miLine9: TMenuItem;
    miLine5: TMenuItem;
    miConcordanceJoin: TMenuItem;
    miDiagramShowGrid: TMenuItem;
    miDiagramAllWordNoBook: TMenuItem;
    miConcordanceRefreshGrid: TMenuItem;
    miLine6: TMenuItem;
    miDiagramAllWordsBook: TMenuItem;
    miLine13: TMenuItem;
    miDiagramShowVal: TMenuItem;
    miDiaZoomIn: TMenuItem;
    miDiaZoomOut: TMenuItem;
    miDiaZoomNormal: TMenuItem;
    miLine14: TMenuItem;
    miLine12: TMenuItem;
    miDiagramSave: TMenuItem;
    miDiagramSingleWordsBook: TMenuItem;
    miDiagram: TMenuItem;
    miConcordanceRemove: TMenuItem;
    miManual: TMenuItem;
    miLineCopyright: TMenuItem;
    miFileNew: TMenuItem;
    miStatisticSave: TMenuItem;
    miStatisticCreate: TMenuItem;
    miLine10: TMenuItem;
    miStatistic: TMenuItem;
    miFileSaveAs: TMenuItem;
    miLine8: TMenuItem;
    miConcordanceDelCont: TMenuItem;
    miLineFileBottom: TMenuItem;
    miFileUpdBookmark: TMenuItem;
    miFileSetBookmark: TMenuItem;
    miConcordanceOpenSkip: TMenuItem;
    miConcordanceSaveSkip: TMenuItem;
    miLine4: TMenuItem;
    miConcodanceShowSelected: TMenuItem;
    miFileSave: TMenuItem;
    miLine7: TMenuItem;
    miConcordanceSaveRep: TMenuItem;
    miCopyrightForm: TMenuItem;
    miCopyright: TMenuItem;
    miLine3: TMenuItem;
    miConcordanceAddSkip: TMenuItem;
    miConcordanceCreate: TMenuItem;
    miFileOpen: TMenuItem;
    miLine0: TMenuItem;
    miFileExit: TMenuItem;
    mmMenu: TMainMenu;
    miConcordance: TMenuItem;
    miFile: TMenuItem;
    meText: TMemo;
    meSkipList: TMemo;
    odOpenDialog: TOpenDialog;
    pnDiagram: TPanel;
    pnListBookmark: TPanel;
    pnTextBottom: TPanel;
    pcMain: TPageControl;
    rgCondFlt: TRadioGroup;
    rgSortBy: TRadioGroup;
    sbDiagram: TScrollBox;
    sdSaveDialog: TSaveDialog;
    sgWordList: TStringGrid;
    sbStatusBar: TStatusBar;
    sgStatistic: TStringGrid;
    spDiagram: TSplitter;
    spText: TSplitter;
    tsDiagram: TTabSheet;
    tsStatistic: TTabSheet;
    tsFile: TTabSheet;
    tsConcordance: TTabSheet;
    udWordsCont: TUpDown;
    procedure apAppPropException(Sender: TObject; E: Exception);
    procedure bnDeselConcClick(Sender: TObject);
    procedure bnFindFirstClick(Sender: TObject);
    procedure bnFindNextClick(Sender: TObject);
    procedure bnReplaceAllClick(Sender: TObject);
    procedure bnReplaceClick(Sender: TObject);
    procedure bnDeselDiagClick(Sender: TObject);
    procedure cbCollatedSortChange(Sender: TObject);
    procedure cbComboDiag2KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure cbComboDiag3KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure cbComboDiag4KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure cbComboDiag5KeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure clDiagBookmarkKeyUp(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure edFindKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure edFltEndExit(Sender: TObject);
    procedure edFltStartExit(Sender: TObject);
    procedure edLocateKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure edReplaceKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure edWordsContChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure lbBookmarksClick(Sender: TObject);
    procedure meSkipListExit(Sender: TObject);
    procedure meTextChange(Sender: TObject);
    procedure meTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure meTextMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
    procedure miConcodanceShowSelectedClick(Sender: TObject);
    procedure miConcordanceAddSkipClick(Sender: TObject);
    procedure miConcordanceCreateClick(Sender: TObject);
    procedure miConcordanceJoinClick(Sender: TObject);
    procedure miFileOpenConcClick(Sender: TObject);
    procedure miConcordanceOpenSkipClick(Sender: TObject);
    procedure miConcordanceRefreshGridClick(Sender: TObject);
    procedure miConcordanceRemoveClick(Sender: TObject);
    procedure miFileSaveConcClick(Sender: TObject);
    procedure miConcordanceSaveRepClick(Sender: TObject);
    procedure miConcordanceDelContClick(Sender: TObject);
    procedure miConcordanceSaveSkipClick(Sender: TObject);
    procedure miCopyrightFormClick(Sender: TObject);
    procedure miDiagramAllWordNoBookClick(Sender: TObject);
    procedure miDiagramShowGridClick(Sender: TObject);
    procedure miDiagramSingleWordsBookClick(Sender: TObject);
    procedure miDiagramAllWordsBookClick(Sender: TObject);
    procedure miDiagramSaveClick(Sender: TObject);
    procedure miDiagramShowValClick(Sender: TObject);
    procedure miDiaZoomInClick(Sender: TObject);
    procedure miDiaZoomNormalClick(Sender: TObject);
    procedure miDiaZoomOutClick(Sender: TObject);
    procedure miFileNewClick(Sender: TObject);
    procedure miManualClick(Sender: TObject);
    procedure miStatisticCreateClick(Sender: TObject);
    procedure miFileSaveAsClick(Sender: TObject);
    procedure miFileSetBookmarkClick(Sender: TObject);
    procedure miFileUpdBookmarkClick(Sender: TObject);
    procedure miFileExitClick(Sender: TObject);
    procedure miFileOpenClick(Sender: TObject);
    procedure miFileSaveClick(Sender: TObject);
    procedure miStatisticSaveClick(Sender: TObject);
    procedure miStatisticSortFreqClick(Sender: TObject);
    procedure miStatisticSortNameClick(Sender: TObject);
    procedure pcMainChange(Sender: TObject);
    procedure pcMainChanging(Sender: TObject; var AllowChange: boolean);
    procedure pmCopyClick(Sender: TObject);
    procedure pmCutClick(Sender: TObject);
    procedure pmPasteClick(Sender: TObject);
    procedure pmSelAllClick(Sender: TObject);
    procedure rgSortBySelectionChanged(Sender: TObject);
    procedure sgStatisticPrepareCanvas(Sender: TObject; aCol, aRow: integer;
      aState: TGridDrawState);
    procedure sgStatisticSelectCell(Sender: TObject; aCol, aRow: integer;
      var CanSelect: boolean);
    procedure sgWordListKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure sgWordListSelection(Sender: TObject; aCol, aRow: integer);
    procedure tsDiagramResize(Sender: TObject);
  private
    procedure CompileGrid(blSort: boolean);
    procedure CreateConc(stStartText: string);
    function CreateContext(GridRow: integer; blSetCursor: boolean): string;
    procedure AddSkipWord;
    function CheckFilter(stWord: string): boolean;
    function CleanField(myField: string): string;
    function CleanXML(stXMLText: string): string;
    procedure CreateDiagramAllWordsBook;
    procedure CreateDiagramAllWordsNoBook;
    procedure CreateDiagramSingleWordsBook;
    procedure CreateStatistic;
    procedure DisableMenuItems;
    procedure EnableMenuItems;
    function FindNextWord(blMessage: boolean): boolean;
    procedure OpenConcordance;
    procedure SaveConcordance;
    procedure SaveReport(stFileName: string);
    procedure CreateBookmarks(blSetCursor: boolean);
    procedure SortWordFreq(var arWordList: array of TRecWordList;
      flField: shortint);
    { private declarations }
  public
    { public declarations }
  end;

var
  fmMain: TfmMain;
  arWordList: array of TRecWordList;
  arWordTextual: array of TRecWordTextual;
  myHomeDir: string;
  myConfigFile: string;
  stFileName: string;
  iWordsUsed: integer = 0;
  iWordsTotal: integer = 0;
  iWordsStartTot: integer = 0;
  ttStart, ttEnd: TTime;
  blSortMod: boolean = False;
  blStopConcordance: boolean = False;
  blConInProc: boolean = False;
  blTextModSave: boolean = False;
  blTextModConc: boolean = False;
  stDiagTitle: string;
  stCpr1: string = '';
  stCpr2: string = '';
  stCpr3: string = '';

resourcestring

  msg001 = 'Operation not correct.';
  msg002 = 'Compact all paragraphs not separated by an empty line, ' +
    'clean the text from double spaces and check the spaces ' +
    'after the punctuation marks?';
  msg003 = 'No recurrences of the text to look for.';
  msg004 = 'No text is selected.';
  msg005 = 'Replace all the recurrences of';
  msg006 = 'with';
  msg007 = 'in the current text?';
  msg008 = 'The word looked for has been replaced';
  msg009 = 'times.';
  msg010 = 'Create a new text and a new concordance ' +
    'closing the existing ones?';
  msg011 = 'The current text has been changed but has not been saved. ' +
    'Reject the changes and open a new text?';
  msg012 = 'It is not possible to open the selected file.';
  msg013 = 'It is not possible to save the selected file.';
  msg014 = 'It is not possible to save the selected file.';
  msg015 = 'There is no word to remove.';
  msg016 = 'No concordance has been created.';
  msg017 = 'No words are selected.';
  msg018 = 'Associate all the recurrences of the selected words ' +
    'to those of the current word';
  msg019 = 'Some changes have been made to the concordance grid. ' +
    'Do you want to reject them and recreate the list of words?';
  msg020 = 'It is not possible to open the skip list.';
  msg021 = 'It is not possible to save the skip list.';
  msg022 = 'There is no recurrence to delete.';
  msg023 = 'There is no selected recurrence to delete.';
  msg024 = 'No concordance has been created.';
  msg025 = 'It is not possible to save the report.';
  msg026 = 'No statistic available.';
  msg027 = 'No statistic available.';
  msg028 = 'There is no statistic to save.';
  msg029 = 'It is not possible to save the selected file.';
  msg030 = 'There is no diagram to save.';
  msg031 = 'It is not possible to save the diagram.';
  msg032 = 'No manual available in the folder of WordStatix. ' +
    'Download it from the website of the software:';
  msg033 = 'A bookmark is not open or closed properly ' +
    'with double square brackets.';
  msg034 = 'A bookmark seems to be too long. ' +
    'Check the text to control that there are ' +
    'no unwanted double square brackets.';
  msg035 = 'A bookmark seems to be incorrect. ' +
    'Check the text to control that there are ' +
    'no unwanted double square brackets.';
  msg036 = 'The bookmark';
  msg037 = 'is used more than once. ' +
    'Check the text and make all its recurrences unique.';
  msg038 = 'Bookmarks cannot contain commas, spaces or be an asterisk. ' +
    'The incorrect bookmarks have been changed accordingly.';
  msg039 = 'No more recurrences of the text to look for.';
  msg040 = 'No text available.';
  msg041 = 'Replace the current concordance with a new one?';
  msg042 = 'The concordance has been created.';
  msg043 = 'Total words (without bookmarks):';
  msg044 = 'Unique words:';
  msg045 = 'Concordance done in';
  msg046 = 'There is no word to add to the skip list.';
  msg047 = 'No text available.';
  msg048 = 'No concordance available.';
  msg049 = 'The text or the settings of the software useful to create ' +
    'the concordance have been changed after the last one was processed. ' +
    'Recreate the concordance to be able to save it along with its text.';
  msg050 = 'It is not possibile to save the concordance.';
  msg051 = 'The current text has been changed but has not been saved. ' +
    'Reject the changes and open a new concordance with its text?';
  msg052 = 'The selected file was not created ' +
    'with WordStatix, and so cannot be opened.';
  msg053 = 'It is not possibile to open the concordance.';
  msg054 = 'No concordance available.';
  msg055 = 'No words are selected in the concordance grid.';
  msg056 = 'No active statistic.';
  msg057 = 'No bookmark available.';
  msg058 = 'At least one bookmark must be selected.';
  msg059 = 'No word selected in Word 1 field.';
  msg060 = 'The fields "Word" must be compiled contiguously.';
  msg061 = 'The same word has been selected twice.';
  msg062 = 'No active statistic.';
  msg063 = 'No bookmark available.';
  msg064 = 'At least one bookmark must be selected.';
  msg065 = 'No active statistic.';
  msg066 = 'No bookmark available.';
  msg067 = 'At least one bookmark must be selected.';
  // Status bar messages
  sbr001 = 'No active concordance.';
  sbr002 = 'The file has been opened.';
  sbr003 = 'The text has been saved.';
  sbr004 = 'The text has been saved.';
  sbr005 = 'The report has been saved.';
  sbr006 = 'The statistic has been saved.';
  sbr007 = 'The diagram has been saved.';
  sbr008 = 'Concordance interrupted.';
  sbr009 = 'Concordance in processing, please wait.';
  sbr010 = 'Concordance in processing, press Ctrl + Shift + H to stop. ' +
    'Analyzed words:';
  sbr011 = 'of';
  sbr012 = 'Unique words found:';
  sbr013 = 'Concordance in processing, press Ctrl + Shift + H to stop. ' +
    'Analyzed words:';
  sbr014 = 'of';
  sbr015 = 'Unique words found:';
  sbr016 = 'Total words (without bookmarks):';
  sbr017 = 'Unique words:';
  sbr018 = 'Concordance done in';
  sbr019 = 'No active concordance.';
  sbr020 = 'Sorting the list of words, press Ctrl + Shift + H to stop.';
  sbr021 = 'Filling the concordance grid, please wait.';
  sbr022 = 'Total words (without bookmarks):';
  sbr023 = 'Unique words:';
  sbr024 = 'Sorting interrupted.';
  sbr025 = 'The concordance has been saved.';
  sbr026 = 'Opening concordance, please wait.';
  sbr027 = 'Total words (without bookmarks):';
  sbr028 = 'Unique words:';
  sbr029 = 'Concordance loaded in';
  // Dialogs and files captions
  dgc001 = 'Open file';
  dgc002 = 'All formats|*.txt;*.odt;*.docx|' +
    'Text file|*.txt|Writer files|*.odt|' + 'Word files|*.docx|All files|*';
  dgc003 = 'Save file';
  dgc004 = 'Text file|*.txt|All files|*';
  dgc005 = 'Open skip list';
  dgc006 = 'Text file|*.txt|All files|*';
  dgc007 = 'Save skip list';
  dgc008 = 'Text file|*.txt|All files|*';
  dgc009 = 'Save report';
  dgc010 = 'Text file|*.txt|HTML file|*.html|All files|*';
  dgc011 = 'Save statistic';
  dgc012 = 'CSV file|*.csv|All files|*';
  dgc013 = 'Save diagram';
  dgc014 = 'JPEG file|*.jpg|PNG file|*.png|All files|*';
  dgc015 = 'Open concordance';
  dgc016 = 'WordStatix files|*.wsx|All files|*';
  dgc017 = 'Title of report';
  dgc018 = 'Insert the title of the report';
  dgc019 = 'Concordance made with WordStatix';
  dgc020 = 'recurrences';
  dgc021 = 'Total';
  dgc022 = 'Title of diagram';
  dgc023 = 'Insert the title of the diagram.';
  srb001 = 'Sort by recurrences';
  srb002 = 'Sort by words';
  man001 = 'manual-wordstatix.pdf';

implementation

{$R *.lfm}

uses unit2, Copyright;

{ TfmMain }

// ********************************************************** //
// ***************** Events of the main form **************** //
// ********************************************************** //

procedure TfmMain.FormCreate(Sender: TObject);
var
  MyIni: TIniFile;
begin
  // Set data on start
  {$ifdef Linux}
  myHomeDir := GetEnvironmentVariable('HOME') + DirectorySeparator +
    '.config' + DirectorySeparator + 'wordstatix' + DirectorySeparator;
  myConfigFile := 'wordstatix';
  {$endif}
  {$ifdef Windows}
  myHomeDir := GetAppConfigDir(False);
  myConfigFile := 'wordstatix.ini';
  fmMain.Color := clWhite;
  lbBookmarks.Color := clWhite;
  clDiagBookmark.Color := clWhite;
  chChart.Color := clWhite;
  chChart.BackColor := clWhite;
  chChart.Title.Font.Color := clBlack;
  chChart.Title.Brush.Color := clWhite;
  {$endif}
  {$ifdef Darwin}
  myHomeDir := GetEnvironmentVariable('HOME') + DirectorySeparator +
    'Library' + DirectorySeparator + 'Preferences' + DirectorySeparator;
  myConfigFile := 'wordstatix.plist';
  miFileOpen.ShortCut := 16463 - 12288;
  miFileSave.ShortCut := 16467 - 12288;
  miFileOpenConc.ShortCut := 24655 - 12288;
  miFileSaveConc.ShortCut := 24659 - 12288;
  miFileSetBookmark.ShortCut := 16466 - 12288;
  miFileUpdBookmark.ShortCut := 16469 - 12288;
  miFileExit.ShortCut := 16465 - 12288;
  miConcordanceCreate.ShortCut := 16462 - 12288;
  miConcodanceShowSelected.ShortCut := 16460 - 12288;
  miConcordanceAddSkip.ShortCut := 24651 - 12288;
  miConcordanceRemove.ShortCut := 16459 - 12288;
  miConcordanceJoin.ShortCut := 16458 - 12288;
  miConcordanceRefreshGrid.ShortCut := 24661 - 12288;
  miConcordanceDelCont.ShortCut := 24644 - 12288;
  miConcordanceSaveRep.ShortCut := 24659 - 12288;
  miStatisticCreate.ShortCut := 16468 - 12288;
  miStatisticSortName.ShortCut := 16473 - 12288;
  miStatisticSortFreq.ShortCut := 24665 - 12288;
  miDiagramAllWordNoBook.ShortCut := 16457 - 12288;
  miDiagramAllWordsBook.ShortCut := 24649 - 12288;
  miDiagramSingleWordsBook.ShortCut := 57417 - 12288;
  miDiaZoomIn.ShortCut := 16571 - 12288;
  miDiaZoomOut.ShortCut := 16573 - 12288;
  miDiaZoomNormal.ShortCut := 16432 - 12288;
  miFile.Caption := #$EF#$A3#$BF;
  miFileCopyright.Visible := True;
  miFileManual.Visible := True;
  miLineManual.Visible := True;
  miCopyright.Visible := False;
  miLineFileBottom.Visible := False;
  miFileExit.Visible := False;
  {$endif}
  sgWordList.FocusRectVisible := False;
  sgStatistic.FocusRectVisible := False;
  lbBookmarks.ScrollWidth := 0;
  clDiagBookmark.ScrollWidth := 0;
  miFileSave.Enabled := False;
  rgSortBy.Items.Clear;
  rgSortBy.Items.Add(srb001);
  rgSortBy.Items.Add(srb002);
  rgSortBy.ItemIndex := 0;
  if DirectoryExists(myHomeDir) = False then
  begin
    CreateDirUTF8(myHomeDir);
  end;
  if FileExistsUTF8(myHomeDir + 'skipwords') then
  begin
    meSkipList.Lines.LoadFromFile(myHomeDir + 'skipwords');
  end;
  if FileExistsUTF8(myHomeDir + myConfigFile) then
  begin
    try
      MyIni := TIniFile.Create(myHomeDir + myConfigFile);
      if MyIni.ReadString('wordstatix', 'maximize', '') = 'true' then
      begin
        fmMain.WindowState := wsMaximized;
      end
      else
      begin
        fmMain.Top := MyIni.ReadInteger('wordstatix', 'top', 0);
        fmMain.Left := MyIni.ReadInteger('wordstatix', 'left', 0);
        if MyIni.ReadInteger('wordstatix', 'width', 0) > 100 then
          fmMain.Width := MyIni.ReadInteger('wordstatix', 'width', 0)
        else
          fmMain.Width := 1000;
        if MyIni.ReadInteger('wordstatix', 'heigth', 0) > 100 then
          fmMain.Height := MyIni.ReadInteger('wordstatix', 'heigth', 0)
        else
          fmMain.Height := 600;
      end;
      if MyIni.ReadInteger('wordstatix', 'sortby', 0) > -1 then
      begin
        rgSortBy.ItemIndex := MyIni.ReadInteger('wordstatix', 'sortby', 0);
      end;
      if MyIni.ReadString('wordstatix', 'wordscont', '') <> '' then
      begin
        edWordsCont.Text := MyIni.ReadString('wordstatix', 'wordscont', '6');
      end;
      if MyIni.ReadString('wordstatix', 'collatesort', '') = 't' then
      begin
        cbCollatedSort.Checked := True;
      end;
      if MyIni.ReadString('wordstatix', 'skipnum', '') = 't' then
      begin
        cbSkipNumbers.Checked := True;
      end;
    finally
      MyIni.Free;
    end;
  end;
end;

procedure TfmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  MyIni: TIniFile;
begin
  // Save data on close
  blStopConcordance := True;
  meSkipList.Lines.SaveToFile(myHomeDir + 'skipwords');
  try
    MyIni := TIniFile.Create(myHomeDir + myConfigFile);
    if fmMain.WindowState = wsMaximized then
    begin
      MyIni.WriteString('wordstatix', 'maximize', 'true');
    end
    else
    begin
      MyIni.WriteString('wordstatix', 'maximize', 'false');
      MyIni.WriteInteger('wordstatix', 'top', fmMain.Top);
      MyIni.WriteInteger('wordstatix', 'left', fmMain.Left);
      MyIni.WriteInteger('wordstatix', 'width', fmMain.Width);
      MyIni.WriteInteger('wordstatix', 'heigth', fmMain.Height);
    end;
    MyIni.WriteString('wordstatix', 'wordscont', edWordsCont.Text);
    MyIni.WriteInteger('wordstatix', 'sortby', rgSortBy.ItemIndex);
    if cbCollatedSort.Checked = True then
    begin
      MyIni.WriteString('wordstatix', 'collatesort', 't');
    end
    else
    begin
      MyIni.WriteString('wordstatix', 'collatesort', 'f');
    end;
    if cbSkipNumbers.Checked = True then
    begin
      MyIni.WriteString('wordstatix', 'skipnum', 't');
    end
    else
    begin
      MyIni.WriteString('wordstatix', 'skipnum', 'f');
    end;
  finally
    MyIni.Free;
  end;
end;

procedure TfmMain.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Select tab
  if ((Key = Ord('1')) and (Shift = [ssCtrl])) then
  begin
    pcMain.ActivePage := tsFile;
    key := 0;
  end
  else if ((Key = Ord('2')) and (Shift = [ssCtrl])) then
  begin
    pcMain.ActivePage := tsConcordance;
    key := 0;
  end
  else if ((Key = Ord('3')) and (Shift = [ssCtrl])) then
  begin
    pcMain.ActivePage := tsStatistic;
    key := 0;
  end
  else if ((Key = Ord('4')) and (Shift = [ssCtrl])) then
  begin
    pcMain.ActivePage := tsDiagram;
    key := 0;
  end;
  // Zoom the diagram
  if ((pcMain.ActivePage = tsDiagram) and (chChart.Visible = True)) then
  begin
    if key = 37 then
    begin
      sbDiagram.HorzScrollBar.Position :=
        sbDiagram.HorzScrollBar.Position - 50;
      key := 0;
    end
    else if key = 39 then
    begin
      sbDiagram.HorzScrollBar.Position :=
        sbDiagram.HorzScrollBar.Position + 50;
      key := 0;
    end;
  end;
  // Stop concordance
  if ((Key = Ord('H')) and (Shift = [ssShift, ssCtrl])) then
  begin
    blStopConcordance := True;
  end;
end;

procedure TfmMain.FormShow(Sender: TObject);
begin
  // Focus on text
  pcMain.ActivePage := tsFile;
  meText.SetFocus;
end;

procedure TfmMain.FormActivate(Sender: TObject);
begin
  // Hack to fix some misalignments in Gnome > 2.16
  // Do not set properties of components with modified anchors
  {$ifdef Linux}
  bnDeselConc.Top := edLocate.Top + edLocate.Height - bnDeselConc.Height;
  {$endif}
end;

procedure TfmMain.apAppPropException(Sender: TObject; E: Exception);
begin
  // Error handling
  MessageDlg(msg001, mtError, [mbOK], 0);
end;

procedure TfmMain.lbBookmarksClick(Sender: TObject);
begin
  // Go to bookmark
  if ((lbBookmarks.Items.Count > 0) and
    (lbBookmarks.ItemIndex > -1)) then
  begin
    meText.SelStart := UTF8Pos('[[' +
      lbBookmarks.Items[lbBookmarks.ItemIndex] + ']]', meText.Text) - 1;
    meText.SelLength := UTF8Length(lbBookmarks.Items[lbBookmarks.ItemIndex]) + 4;
    meText.SetFocus;
  end;
end;

procedure TfmMain.pcMainChange(Sender: TObject);
begin
  // Selecting the tsConcordance the sdGrid do not accept focus, so...
  if pcMain.ActivePage = tsConcordance then
    tsConcordance.SetFocus;
end;

procedure TfmMain.pcMainChanging(Sender: TObject; var AllowChange: boolean);
begin
  // Prevent change tab during concordance
  if blConInProc = True then
  begin
    AllowChange := False;
  end
  else
  begin
    AllowChange := True;
  end;
end;

procedure TfmMain.pmCutClick(Sender: TObject);
begin
  // Cut
  meText.CutToClipboard;
end;

procedure TfmMain.pmCopyClick(Sender: TObject);
begin
  // Copy
  meText.CopyToClipboard;
end;

procedure TfmMain.pmPasteClick(Sender: TObject);
begin
  // Paste
  meText.PasteFromClipboard;
end;

procedure TfmMain.pmSelAllClick(Sender: TObject);
begin
  // Select all
  meText.SelectAll;
end;

procedure TfmMain.rgSortBySelectionChanged(Sender: TObject);
begin
  // Changed settings for concordance
  // also for rgCondFlt
  blTextModConc := True;
end;

procedure TfmMain.cbCollatedSortChange(Sender: TObject);
begin
  // Changed settings for concordance
  // also for all other Edit and Memo components
  blTextModConc := True;
end;

procedure TfmMain.sgWordListKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Select and deselect the word
  if key = 32 then
  begin
    sgWordList.Col := 0;
    if sgWordList.Cells[2, sgWordList.Row] = '0' then
    begin
      sgWordList.Cells[2, sgWordList.Row] := '1';
    end
    else
    begin
      sgWordList.Cells[2, sgWordList.Row] := '0';
    end;
    if sgWordList.Row < sgWordList.RowCount - 1 then
    begin
      sgWordList.Row := sgWordList.Row + 1;
    end;
  end;
end;

procedure TfmMain.bnDeselConcClick(Sender: TObject);
var
  i: integer;
begin
  // Select and deselect all words
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.Cells[2, 1] = '1' then
    begin
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        sgWordList.Cells[2, i] := '0';
      end;
    end
    else
    begin
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        sgWordList.Cells[2, i] := '1';
      end;
    end;
  end;
end;

procedure TfmMain.sgWordListSelection(Sender: TObject; aCol, aRow: integer);
begin
  // Create the list of context
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.RowHeights[sgWordList.Row] > 0 then
    begin
      lsContext.Clear;
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
      {$ifdef Windows}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
      {$endif}
      lsContext.ItemIndex := 0;
    end;
  end;
end;

procedure TfmMain.sgStatisticPrepareCanvas(Sender: TObject;
  aCol, aRow: integer; aState: TGridDrawState);
begin
  // Font of the last row and second col in Statistic
  if aRow = sgStatistic.RowCount - 1 then
  begin
    sgStatistic.Canvas.Font.Style := [fsBold];
    sgStatistic.Canvas.Brush.Color := clBtnFace;
  end
  else
  if aCol = 1 then
  begin
    sgStatistic.Canvas.Font.Style := [fsBold];
  end;
end;

procedure TfmMain.sgStatisticSelectCell(Sender: TObject; aCol, aRow: integer;
  var CanSelect: boolean);
begin
  // Avoid select last row
  if aRow = sgStatistic.RowCount - 1 then
  begin
    CanSelect := False;
  end;
end;

procedure TfmMain.edWordsContChange(Sender: TObject);
begin
  // Create the list of context
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.RowHeights[sgWordList.Row] > 0 then
    begin
      lsContext.Clear;
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
      {$ifdef Windows}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
      {$endif}
      lsContext.ItemIndex := 0;
    end;
  end;
end;

procedure TfmMain.meSkipListExit(Sender: TObject);
begin
  // Clean skip words on exit
  meSkipList.Text := CleanField(meSkipList.Text);
end;

procedure TfmMain.meTextChange(Sender: TObject);
begin
  // Set the text as modified
  blTextModSave := True;
  miFileSave.Enabled := blTextModSave;
  blTextModConc := True;
end;

procedure TfmMain.edFltStartExit(Sender: TObject);
begin
  // Clean filter start with on exit
  edFltStart.Text := CleanField(edFltStart.Text);
end;

procedure TfmMain.edLocateKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
var
  i: integer;
begin
  // Locate word in concordance
  if edLocate.Text <> '' then
  begin
    if ((key = 13) and (Shift = [ssCtrl])) then
    begin
      if sgWordList.Row < sgWordList.RowCount - 1 then
      begin
        for i := sgWordList.Row + 1 to sgWordList.RowCount - 1 do
        begin
          if UTF8LowerCase(UTF8Copy(sgWordList.Cells[0, i], 1,
            UTF8Length(edLocate.Text))) = UTF8LowerCase(edLocate.Text) then
          begin
            if sgWordList.RowHeights[i] > 0 then
            begin
              sgWordList.Row := i;
            end;
            Break;
          end;
        end;
      end;
    end
    else
    if key = 13 then
    begin
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        if UTF8LowerCase(UTF8Copy(sgWordList.Cells[0, i], 1,
          UTF8Length(edLocate.Text))) = UTF8LowerCase(edLocate.Text) then
        begin
          if sgWordList.RowHeights[i] > 0 then
          begin
            sgWordList.Row := i;
          end;
          Break;
        end;
      end;
    end;
  end;
end;

procedure TfmMain.edFindKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
var
  stClip: string;
begin
  // Insert indivisible space
  if ((Key = Ord(' ')) and (Shift = [ssCtrl])) then
  begin
    stClip := Clipboard.AsText;
    Clipboard.AsText := ' ';
    edFind.PasteFromClipboard;
    Clipboard.AsText := stClip;
  end;
end;

procedure TfmMain.edReplaceKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
var
  stClip: string;
begin
  // Insert indivisible space
  if ((Key = Ord(' ')) and (Shift = [ssCtrl])) then
  begin
    stClip := Clipboard.AsText;
    Clipboard.AsText := ' ';
    edReplace.PasteFromClipboard;
    Clipboard.AsText := stClip;
  end;
end;

procedure TfmMain.edFltEndExit(Sender: TObject);
begin
  // Clean filter end on exit
  edFltEnd.Text := CleanField(edFltEnd.Text);
end;

procedure TfmMain.meTextMouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
begin
  // Change Zoom with mouse weel
  {$ifdef Linux}
  if Shift = [ssCtrl] then
  begin
    if WheelDelta > 0 then
    begin
      if meText.Font.Size < 48 then
        meText.Font.Size := meText.Font.Size + 1;
    end
    else
    begin
      if meText.Font.Size > 6 then
        meText.Font.Size := meText.Font.Size - 1;
    end;
  end;
  {$endif}
  {$ifdef Windows}
  if Shift = [ssCtrl] then
  begin
    if WheelDelta > 0 then
    begin
      if meText.Font.Size < 48 then
        meText.Font.Size := meText.Font.Size + 1;
    end
    else
    begin
      if meText.Font.Size > 6 then
        meText.Font.Size := meText.Font.Size - 1;
    end;
  end;
  {$endif}
end;

procedure TfmMain.meTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
var
  stClip: string;
begin
  // Change Zoom with keys
  {$ifdef Linux}
  if ((Key = 187) and (Shift = [ssShift, ssCtrl])) then
  begin
    if meText.Font.Size < 48 then
      meText.Font.Size := meText.Font.Size + 1;
  end
  else if ((Key = 189) and (Shift = [ssShift, ssCtrl])) then
  begin
    if meText.Font.Size > 6 then
      meText.Font.Size := meText.Font.Size - 1;
  end
  else if ((Key = Ord('0')) and (Shift = [ssCtrl])) then
  begin
    meText.Font.Size := 14;
  end
  else
  {$endif}
  {$ifdef Windows}
  if ((Key = 187) and (Shift = [ssShift, ssCtrl])) then
  begin
    if meText.Font.Size < 48 then
      meText.Font.Size := meText.Font.Size + 1;
  end
  else if ((Key = 189) and (Shift = [ssShift, ssCtrl])) then
  begin
    if meText.Font.Size > 6 then
      meText.Font.Size := meText.Font.Size - 1;
  end
  else if ((Key = Ord('0')) and (Shift = [ssCtrl])) then
  begin
    meText.Font.Size := 14;
  end
  else
  {$endif}
  // Insert indivisible space
  if ((Key = Ord(' ')) and (Shift = [ssCtrl])) then
  begin
    stClip := Clipboard.AsText;
    Clipboard.AsText := ' ';
    meText.PasteFromClipboard;
    Clipboard.AsText := stClip;
  end
  // Clean text
  else if ((Key = Ord('P')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if meText.Text = '' then
    begin
      Abort;
    end;
    if MessageDlg(msg002, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
      Abort;
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      meText.Text := StringReplace(meText.Text, '[[', #2, [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '[', ' [', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '(', ' (', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, #2, '[[', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '[[', ' [[', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '.', '. ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, ',', ', ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, ';', '; ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, ':', ': ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '!', '! ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, '?', '? ', [rfReplaceAll]);
      while UTF8Pos(LineEnding + LineEnding + LineEnding, meText.Text) > 0 do
      begin
        meText.Text := StringReplace(meText.Text, LineEnding +
          LineEnding + LineEnding, LineEnding + LineEnding, [rfReplaceAll]);
      end;
      meText.Text := StringReplace(meText.Text, LineEnding + LineEnding,
        #2, [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, LineEnding, ' ', [rfReplaceAll]);
      meText.Text := StringReplace(meText.Text, #2, LineEnding +
        LineEnding, [rfReplaceAll]);
      while UTF8Pos('  ', meText.Text) > 0 do
      begin
        meText.Text := StringReplace(meText.Text, '  ', ' ', [rfReplaceAll]);
      end;
      meText.Text := StringReplace(meText.Text, LineEnding + ' [[',
        LineEnding + '[[', [rfReplaceAll]);
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

procedure TfmMain.bnFindFirstClick(Sender: TObject);
begin
  // Find first
  if edFind.Text <> '' then
  begin
    if UTF8Pos(UTF8LowerCase(edFind.Text), UTF8LowerCase(meText.Text), 1) > 0 then
    begin
      meText.SelStart := UTF8Pos(UTF8LowerCase(edFind.Text),
        UTF8LowerCase(meText.Text)) - 1;
      meText.SelLength := UTF8Length(edFind.Text);
      meText.SetFocus;
    end
    else
    begin
      MessageDlg(msg003,
        mtInformation, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.bnFindNextClick(Sender: TObject);
begin
  // Find next
  FindNextWord(True);
end;

procedure TfmMain.bnReplaceClick(Sender: TObject);
  var iPos: integer;
begin
  // Replace selection
  if meText.SelLength = 0 then
  begin
    MessageDlg(msg004, mtWarning, [mbOK], 0);
    Abort;
  end;
  if edReplace.Text <> '' then
  begin
    iPos := meText.SelStart;
    meText.Text := StringReplace(meText.Text, meText.SelText, edReplace.Text, []);
    meText.SelStart := iPos + UTF8Length(edReplace.Text) + 1;
    if edFind.Text <> '' then
    begin
      bnFindNextClick(nil);
    end;
  end;
end;

procedure TfmMain.bnReplaceAllClick(Sender: TObject);
var
  stText: string;
  i, iCount: integer;
begin
  // Replace all
  if ((edFind.Text <> '') and (edReplace.Text <> '')) then
  begin
    if MessageDlg(msg005 + ' "' + edFind.Text + '" ' + msg006 + ' "' +
      edReplace.Text + '" ' + msg007, mtConfirmation, [mbOK, mbCancel], 0) =
      mrCancel then
    begin
      Abort;
    end;
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    stText := meText.Text;
    i := 1;
    iCount := 0;
    try
      while UTF8Pos(UTF8LowerCase(edFind.Text), UTF8LowerCase(stText), i) > 0 do
      begin
        i := UTF8Pos(UTF8LowerCase(edFind.Text), UTF8LowerCase(stText), i);
        stText := UTF8Copy(stText, 1, i - 1) + edReplace.Text +
          UTF8Copy(stText, i + UTF8Length(edFind.Text), UTF8Length(stText));
        i := i + UTF8Length(edReplace.Text);
        Inc(iCount);
        Application.ProcessMessages;
      end;
      meText.Text := stText;
    finally
      Screen.Cursor := crDefault;
    end;
    MessageDlg(msg008 + ' ' + IntToStr(iCount) + ' ' + msg009,
      mtInformation, [mbOK], 0);
  end;
end;

procedure TfmMain.tsDiagramResize(Sender: TObject);
begin
  // Resize the diagram
  chChart.Width := sbDiagram.Width - 3;
  chChart.Height := sbDiagram.Height - 10;
end;

procedure TfmMain.bnDeselDiagClick(Sender: TObject);
begin
  // Select and deselect all bookmarks
  if clDiagBookmark.Count > 0 then
  begin
    if clDiagBookmark.Checked[0] = False then
    begin
      clDiagBookmark.CheckAll(cbChecked, False, False);
    end
    else
    begin
      clDiagBookmark.CheckAll(cbUnchecked, False, False);
    end;
  end;
end;

procedure TfmMain.cbComboDiag2KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Clear the word field
  if ((Key = 46) or (key = 8)) then
    cbComboDiag2.Text := '';
end;

procedure TfmMain.cbComboDiag3KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Clear the word field
  if ((Key = 46) or (key = 8)) then
    cbComboDiag3.Text := '';
end;

procedure TfmMain.cbComboDiag4KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Clear the word field
  if ((Key = 46) or (key = 8)) then
    cbComboDiag4.Text := '';
end;

procedure TfmMain.cbComboDiag5KeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Clear the word field
  if ((Key = 46) or (key = 8)) then
    cbComboDiag5.Text := '';
end;

procedure TfmMain.clDiagBookmarkKeyUp(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  // Move on after space bar
  if key = 32 then
  begin
    if clDiagBookmark.ItemIndex < clDiagBookmark.Items.Count - 1 then
    begin
      clDiagBookmark.ItemIndex := clDiagBookmark.ItemIndex + 1;
    end;
  end;
end;

// ********************************************************** //
// ****************** Events of menu items ****************** //
// ********************************************************** //

procedure TfmMain.miFileNewClick(Sender: TObject);
begin
  // New
  if meText.Text <> '' then
  begin
    if MessageDlg(msg010, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    begin
      Abort;
    end;
  end;
  SetLength(arWordList, 0);
  SetLength(arWordTextual, 0);
  iWordsUsed := 0;
  iWordsTotal := 0;
  meText.Clear;
  lbBookmarks.Clear;
  edFind.Clear;
  edReplace.Clear;
  sgWordList.RowCount := 1;
  lsContext.Clear;
  edFltStart.Clear;
  edFltEnd.Clear;
  sgStatistic.RowCount := 0;
  sgStatistic.ColCount := 0;
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  cbComboDiag1.Clear;
  cbComboDiag2.Clear;
  cbComboDiag3.Clear;
  cbComboDiag4.Clear;
  cbComboDiag5.Clear;
  clDiagBookmark.Clear;
  stFileName := '';
  fmMain.Caption := 'WordStatix';
  miFileSave.Enabled := False;
  pcMain.ActivePage := tsFile;
  sbStatusBar.SimpleText := sbr001;
end;

procedure TfmMain.miFileOpenClick(Sender: TObject);
var
  myZip: TUnZipper;
  myList, stFileOrig: TStringList;
  i: integer;
begin
  // Open file
  pcMain.ActivePage := tsFile;
  if ((meText.Text <> '') and (blTextModSave = True)) then
  begin
    if MessageDlg(msg011, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    begin
      Abort;
    end;
  end;
  odOpenDialog.Title := dgc001;
  odOpenDialog.Filter := dgc002;
  odOpenDialog.DefaultExt := '.txt';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute then
    try
      SetLength(arWordList, 0);
      SetLength(arWordTextual, 0);
      iWordsUsed := 0;
      iWordsTotal := 0;
      meText.Clear;
      lbBookmarks.Clear;
      edFind.Clear;
      edReplace.Clear;
      sgWordList.RowCount := 1;
      lsContext.Clear;
      edFltStart.Clear;
      edFltEnd.Clear;
      sgStatistic.RowCount := 0;
      sgStatistic.ColCount := 0;
      chChart.Visible := False;
      chChartBarSeries1.Active := False;
      chChartBarSeries2.Active := False;
      chChartBarSeries3.Active := False;
      chChartBarSeries4.Active := False;
      chChartBarSeries5.Active := False;
      cbComboDiag1.Clear;
      cbComboDiag2.Clear;
      cbComboDiag3.Clear;
      cbComboDiag4.Clear;
      cbComboDiag5.Clear;
      clDiagBookmark.Clear;
      if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '.txt' then
      begin
        meText.Lines.LoadFromFile(odOpenDialog.FileName);
        stFileName := odOpenDialog.FileName;
        blTextModSave := False;
        miFileSave.Enabled := blTextModSave;
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end
      else if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '' then
      begin
        meText.Lines.LoadFromFile(odOpenDialog.FileName);
        stFileName := odOpenDialog.FileName;
        blTextModSave := False;
        miFileSave.Enabled := blTextModSave;
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end
      else if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '.odt' then
      begin
        try
          Screen.Cursor := crHourGlass;
          Application.ProcessMessages;
          myZip := TUnZipper.Create;
          myList := TStringList.Create;
          stFileOrig := TStringList.Create;
          myList.Add('content.xml');
          myZip.OutputPath := GetTempDir;
          myZip.FileName := odOpenDialog.FileName;
          myZip.UnZipFiles(myList);
          stFileOrig.LoadFromFile(GetTempDir + DirectorySeparator + 'content.xml');
          stFileOrig.Text := StringReplace(stFileOrig.Text,
            '<text:note-citation>', ' [', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text,
            '</text:p></text:note-body></text:note>', ']', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '<text:note-body>',
            ': ', [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '</text:h>',
            LineEnding + LineEnding, [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '</text:p>',
            LineEnding, [rfReplaceAll]);
          stFileOrig.Text := StringReplace(stFileOrig.Text, '&apos;',
            #39, [rfReplaceAll]);
          meText.Text := CleanXML(stFileOrig.Text);
          blTextModSave := False;
          miFileSave.Enabled := blTextModSave;
          DeleteFileUTF8(GetTempDir + DirectorySeparator + 'content.xml');
        finally
          myZip.Free;
          myList.Free;
          stFileOrig.Free;
          Screen.Cursor := crDefault;
        end;
        stFileName := ExtractFileNameWithoutExt(odOpenDialog.FileName) + '.txt';
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end
      else
      if UTF8LowerCase(ExtractFileExt(odOpenDialog.FileName)) = '.docx' then
      begin
        try
          Screen.Cursor := crHourGlass;
          Application.ProcessMessages;
          myZip := TUnZipper.Create;
          myList := TStringList.Create;
          stFileOrig := TStringList.Create;
          myZip.OutputPath := GetTempDir;
          myZip.FileName := odOpenDialog.FileName;
          myZip.UnZipAllFiles;
          if FileExistsUTF8(GetTempDir + 'word/document.xml') = True then
          begin
            stFileOrig.LoadFromFile(GetTempDir + 'word/document.xml');
            stFileOrig.Text :=
              StringReplace(stFileOrig.Text, '</w:p>', LineEnding, [rfReplaceAll]);
            i := 0;
            while Pos('<w:footnoteReference w:id="', stFileOrig.Text) > 0 do
            begin
              Inc(i);
              stFileOrig.Text :=
                StringReplace(stFileOrig.Text, '<w:footnoteReference w:id="',
                ' [' + IntToStr(i) + ']<', []);
            end;
            i := 0;
            while Pos('<w:endnoteReference w:id="', stFileOrig.Text) > 0 do
            begin
              Inc(i);
              stFileOrig.Text :=
                StringReplace(stFileOrig.Text, '<w:endnoteReference w:id="',
                ' [' + IntToStr(i) + ']<', []);
            end;
            meText.Text := meText.Text + CleanXML(stFileOrig.Text) + LineEnding;
          end;
          if FileExistsUTF8(GetTempDir + 'word/footnotes.xml') = True then
          begin
            stFileOrig.LoadFromFile(GetTempDir + 'word/footnotes.xml');
            stFileOrig.Text :=
              StringReplace(stFileOrig.Text, '</w:p>', LineEnding, [rfReplaceAll]);
            i := 0;
            while Pos('<w:footnoteRef/>', stFileOrig.Text) > 0 do
            begin
              Inc(i);
              stFileOrig.Text :=
                StringReplace(stFileOrig.Text, '<w:footnoteRef/>',
                '>[' + IntToStr(i) + '] ', []);
            end;
            meText.Text := meText.Text + CleanXML(stFileOrig.Text) + LineEnding;
          end;
          if FileExistsUTF8(GetTempDir + 'word/endnotes.xml') = True then
          begin
            stFileOrig.LoadFromFile(GetTempDir + 'word/endnotes.xml');
            stFileOrig.Text :=
              StringReplace(stFileOrig.Text, '</w:p>', LineEnding, [rfReplaceAll]);
            i := 0;
            while Pos('<w:endnoteRef/>', stFileOrig.Text) > 0 do
            begin
              Inc(i);
              stFileOrig.Text :=
                StringReplace(stFileOrig.Text, '<w:endnoteRef/>',
                '>[endnote: ' + IntToStr(i) + '] ', []);
            end;
            meText.Text := meText.Text + CleanXML(stFileOrig.Text) + LineEnding;
            blTextModSave := False;
            miFileSave.Enabled := blTextModSave;
          end;
          if DirectoryExistsUTF8(GetTempDir + 'word') = True then
          begin
            DeleteDirectory(GetTempDir + 'word', False);
          end;
        finally
          myZip.Free;
          myList.Free;
          stFileOrig.Free;
          Screen.Cursor := crDefault;
        end;
        stFileName := ExtractFileNameWithoutExt(odOpenDialog.FileName) + '.txt';
        fmMain.Caption := 'WordStatix - ' + stFileName;
      end;
      CreateBookmarks(True);
      sbStatusBar.SimpleText := sbr002;
    except
      MessageDlg(msg012,
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miFileSaveClick(Sender: TObject);
begin
  // Save file
  pcMain.ActivePage := tsFile;
  if stFileName <> '' then
    try
      meText.Lines.SaveToFile(stFileName);
      sbStatusBar.SimpleText := sbr003;
      blTextModSave := False;
      miFileSave.Enabled := blTextModSave;
      fmMain.Caption := 'WordStatix - ' + stFileName;
    except
      MessageDlg(msg013,
        mtWarning, [mbOK], 0);
    end
  else
  begin
    miFileSaveAsClick(nil);
  end;
end;

procedure TfmMain.miFileSaveAsClick(Sender: TObject);
begin
  // Salve as
  pcMain.ActivePage := tsFile;
  sdSaveDialog.Title := dgc003;
  sdSaveDialog.Filter := dgc004;
  sdSaveDialog.DefaultExt := '.txt';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      stFileName := sdSaveDialog.FileName;
      meText.Lines.SaveToFile(stFileName);
      sbStatusBar.SimpleText := sbr004;
      blTextModSave := False;
      miFileSave.Enabled := blTextModSave;
      fmMain.Caption := 'WordStatix - ' + stFileName;
    except
      MessageDlg(msg014,
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miFileOpenConcClick(Sender: TObject);
begin
  // Open concordance
  OpenConcordance;
end;

procedure TfmMain.miFileSaveConcClick(Sender: TObject);
begin
  // Save concordance
  SaveConcordance;
end;

procedure TfmMain.miFileSetBookmarkClick(Sender: TObject);
  var iPos: integer;
begin
  // Set bookmark
  pcMain.ActivePage := tsFile;
  iPos := meText.SelStart;
  if meText.SelText = '' then
  begin
    meText.Text := UTF8Copy(meText.Text, 1, iPos) + '[[]] ' +
      UTF8Copy(meText.Text, iPos + 1, UTF8Length(meText.Text));
    meText.SelStart := iPos + 2;
  end
  else
  begin
    meText.Text := UTF8Copy(meText.Text, 1, iPos) + '[[' + meText.SelText + ']] ' +
      UTF8Copy(meText.Text, iPos + UTF8Length(meText.SelText) + 1,
      UTF8Length(meText.Text));
    meText.SelStart := iPos + 2;
  end;
  CreateBookmarks(True);
end;

procedure TfmMain.miFileUpdBookmarkClick(Sender: TObject);
begin
  // Update bookmarks
  pcMain.ActivePage := tsFile;
  CreateBookmarks(True);
end;

procedure TfmMain.miFileExitClick(Sender: TObject);
begin
  // Quit
  pcMain.ActivePage := tsFile;
  Close;
end;

procedure TfmMain.miConcordanceCreateClick(Sender: TObject);
begin
  // Compile concordance
  pcMain.ActivePage := tsConcordance;
  if sgWordList.Visible = True then
    sgWordList.SetFocus;
  CreateConc(meText.Text);
  if sgWordList.Visible = True then
    sgWordList.SetFocus;
end;

procedure TfmMain.miConcodanceShowSelectedClick(Sender: TObject);
var
  i, iFirstVis: integer;
begin
  // Show only selected words
  pcMain.ActivePage := tsConcordance;
  if sgWordList.RowCount < 2 then
  begin
    Abort;
  end;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  iFirstVis := 0;
  for i := 1 to sgWordList.RowCount - 1 do
  begin
    try
      if miConcodanceShowSelected.Checked = True then
      begin
        if sgWordList.Cells[2, i] = '0' then
        begin
          sgWordList.RowHeights[i] := 0;
        end
        else
        begin
          if iFirstVis = 0 then
          begin
            iFirstVis := i;
          end;
        end;
      end
      else
      begin
        sgWordList.RowHeights[i] := sgWordList.DefaultRowHeight;
        if iFirstVis = 0 then
        begin
          iFirstVis := i;
        end;
      end;
      sgWordList.SetFocus;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
  if iFirstVis > 0 then
  begin
    sgWordList.Row := iFirstVis;
  end;
  if iFirstVis = 0 then
  begin
    lsContext.Clear;
  end
  else
  begin
    lsContext.Items.Text := CreateContext(sgWordList.Row, True);
    {$ifdef Windows}
    // Due to a bug in Windows...
    if lsContext.Items[0] = '' then
      lsContext.Items.Delete(0);
    {$endif}
    lsContext.ItemIndex := 0;
  end;
end;

procedure TfmMain.miConcordanceAddSkipClick(Sender: TObject);
begin
  // Add a word to skip
  pcMain.ActivePage := tsConcordance;
  AddSkipWord;
end;

procedure TfmMain.miConcordanceRemoveClick(Sender: TObject);
begin
  // Remove current word
  pcMain.ActivePage := tsConcordance;
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.RowHeights[sgWordList.Row] > 0 then
    begin
      if sgWordList.Cells[0, sgWordList.Row] <> '' then
      begin
        sgWordList.DeleteRow(sgWordList.Row);
        blSortMod := True;
      end;
    end;
    lsContext.Clear;
    if sgWordList.RowCount > 1 then
    begin
      if sgWordList.RowHeights[sgWordList.Row] > 0 then
      begin
        lsContext.Items.Text := CreateContext(sgWordList.Row, True);
        {$ifdef Windows}
        // Due to a bug in Windows...
        if lsContext.Items[0] = '' then
          lsContext.Items.Delete(0);
        {$endif}
        lsContext.ItemIndex := 0;
      end;
    end;
  end
  else
  begin
    MessageDlg(msg015, mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.miConcordanceJoinClick(Sender: TObject);
var
  i, iPos: integer;
  blSelection: boolean;
begin
  // Join selected words with current word
  if sgWordList.RowHeights[sgWordList.Row] = 0 then
  begin
    Abort;
  end;
  if sgWordList.RowCount < 2 then
  begin
    MessageDlg(msg016, mtWarning, [mbOK], 0);
    Abort;
  end;
  blSelection := False;
  sgWordList.Cells[2, sgWordList.Row] := '0';
  for i := 1 to sgWordList.RowCount - 1 do
  begin
    if sgWordList.Cells[2, i] = '1' then
    begin
      blSelection := True;
      Break;
    end;
  end;
  if blSelection = False then
  begin
    MessageDlg(msg017, mtWarning, [mbOK], 0);
    Abort;
  end;
  if MessageDlg(msg018 + ' "' + sgWordList.Cells[0, sgWordList.Row] +
    '"?', mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
  begin
    Abort;
  end;
  for i := 1 to sgWordList.RowCount - 1 do
  begin
    if sgWordList.Cells[2, i] = '1' then
    begin
      sgWordList.Cells[1, sgWordList.Row] :=
        IntToStr(StrToInt(sgWordList.Cells[1, sgWordList.Row]) +
        StrToInt(sgWordList.Cells[1, i]));
      sgWordList.Cells[3, sgWordList.Row] :=
        sgWordList.Cells[3, sgWordList.Row] + sgWordList.Cells[3, i];
      sgWordList.Cells[4, sgWordList.Row] :=
        sgWordList.Cells[4, sgWordList.Row] + sgWordList.Cells[4, i];
    end;
  end;
  i := 1;
  iPos := sgWordList.Row;
  while i < sgWordList.RowCount do
  begin
    if sgWordList.Cells[2, i] = '1' then
    begin
      sgWordList.DeleteRow(i);
      if i < iPos then
      begin
        iPos := iPos - 1;
      end;
    end
    else
    begin
      Inc(i);
    end;
  end;
  blSortMod := True;
  sgWordList.Row := iPos;
  lsContext.Clear;
  lsContext.Items.Text := CreateContext(sgWordList.Row, True);
  {$ifdef Windows}
  // Due to a bug in Windows...
  if lsContext.Items[0] = '' then
    lsContext.Items.Delete(0);
  {$endif}
  lsContext.ItemIndex := 0;
end;

procedure TfmMain.miConcordanceRefreshGridClick(Sender: TObject);
begin
  // Refresh concordance grid
  pcMain.ActivePage := tsConcordance;
  if blSortMod = True then
  begin
    if MessageDlg(msg019, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    begin
      Abort;
    end;
  end;
  CompileGrid(True);
end;

procedure TfmMain.miConcordanceOpenSkipClick(Sender: TObject);
begin
  // Open skip list
  pcMain.ActivePage := tsConcordance;
  odOpenDialog.Title := dgc005;
  odOpenDialog.Filter := dgc006;
  odOpenDialog.DefaultExt := '.txt';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute then
    try
      meSkipList.Lines.LoadFromFile(odOpenDialog.FileName);
    except
      MessageDlg(msg020,
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miConcordanceSaveSkipClick(Sender: TObject);
begin
  // Save skip list
  pcMain.ActivePage := tsConcordance;
  sdSaveDialog.Title := dgc007;
  sdSaveDialog.Filter := dgc008;
  sdSaveDialog.DefaultExt := '.txt';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      meSkipList.Lines.SaveToFile(sdSaveDialog.FileName);
    except
      MessageDlg(msg021,
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miConcordanceDelContClick(Sender: TObject);
var
  slContList: TStringList;
  iIDList: integer;
begin
  // Selete selected recurrence
  if lsContext.Items.Count < 1 then
  begin
    MessageDlg(msg022, mtWarning, [mbOK], 0);
    Abort;
  end;
  if lsContext.ItemIndex < 0 then
  begin
    MessageDlg(msg023, mtWarning, [mbOK], 0);
    Abort;
  end;
  blSortMod := True;
  if lsContext.Items.Count = 1 then
  begin
    sgWordList.DeleteRow(sgWordList.Row);
    if sgWordList.RowCount > 1 then
    begin
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
      {$ifdef Windows}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
      {$endif}
      lsContext.ItemIndex := 0;
    end
    else
    begin
      lsContext.Clear;
    end;
  end
  else
    try
      slContList := TStringList.Create;
      slContList.CommaText := sgWordList.Cells[3, sgWordList.Row];
      slContList.Delete(lsContext.ItemIndex);
      sgWordList.Cells[3, sgWordList.Row] := slContList.CommaText;
      slContList.CommaText := sgWordList.Cells[4, sgWordList.Row];
      slContList.Delete(lsContext.ItemIndex);
      sgWordList.Cells[4, sgWordList.Row] := slContList.CommaText;
      sgWordList.Cells[1, sgWordList.Row] :=
        IntToStr(StrToInt(sgWordList.Cells[1, sgWordList.Row]) - 1);
      iIDList := lsContext.ItemIndex;
      // Sometimes it might gives error, so
      if lsContext.Items.Count > 0 then
      begin
        lsContext.Items.Delete(lsContext.ItemIndex);
      end;
      if iIDList < lsContext.Items.Count - 1 then
      begin
        lsContext.ItemIndex := iIDList;
      end
      else
      begin
        lsContext.ItemIndex := lsContext.Items.Count - 1;
      end;
    finally
      slContList.Free;
    end;
end;

procedure TfmMain.miConcordanceSaveRepClick(Sender: TObject);
begin
  // Export report to file
  pcMain.ActivePage := tsConcordance;
  if sgWordList.RowCount < 2 then
  begin
    MessageDlg(msg024, mtWarning, [mbOK], 0);
    Abort;
  end;
  sdSaveDialog.Title := dgc009;
  sdSaveDialog.Filter := dgc010;
  sdSaveDialog.DefaultExt := '.txt';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      SaveReport(sdSaveDialog.FileName);
      sbStatusBar.SimpleText := sbr005;
    except
      MessageDlg(msg025,
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miStatisticCreateClick(Sender: TObject);
begin
  // Create statistic
  pcMain.ActivePage := tsStatistic;
  CreateStatistic;
end;

procedure TfmMain.miStatisticSortNameClick(Sender: TObject);
var
  iPos1, iPos2, iCol: integer;
  slRow: TStringList;
begin
  // Sort by name the statistic
  pcMain.ActivePage := tsStatistic;
  if sgStatistic.RowCount = 0 then
  begin
    MessageDlg(msg026, mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    slRow := TStringList.Create;
    for iPos1 := 1 to sgStatistic.RowCount - 2 do
    begin
      for iPos2 := 1 to sgStatistic.RowCount - 3 do
      begin
        if cbCollatedSort.Checked = True then
        begin
          if UTF8CompareStrCollated(sgStatistic.Cells[0, iPos2],
            sgStatistic.Cells[0, iPos1]) > 0 then
          begin
            slRow.Clear;
            for iCol := 0 to sgStatistic.ColCount - 1 do
            begin
              slRow.Add(sgStatistic.Cells[iCol, iPos1]);
            end;
            sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
            for iCol := 0 to slRow.Count - 1 do
            begin
              sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
            end;
          end;
        end
        else
        begin
          if UTF8CompareStr(sgStatistic.Cells[0, iPos2],
            sgStatistic.Cells[0, iPos1]) > 0 then
          begin
            slRow.Clear;
            for iCol := 0 to sgStatistic.ColCount - 1 do
            begin
              slRow.Add(sgStatistic.Cells[iCol, iPos1]);
            end;
            sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
            for iCol := 0 to slRow.Count - 1 do
            begin
              sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
            end;
          end;
        end;
      end;
    end;
  finally
    slRow.Free;
  end;
end;

procedure TfmMain.miStatisticSortFreqClick(Sender: TObject);
var
  iPos1, iPos2, iCol: integer;
  slRow: TStringList;
begin
  // Sort by frequency the statistic
  pcMain.ActivePage := tsStatistic;
  if sgStatistic.RowCount = 0 then
  begin
    MessageDlg(msg027, mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    slRow := TStringList.Create;
    for iPos1 := 1 to sgStatistic.RowCount - 2 do
    begin
      for iPos2 := 1 to sgStatistic.RowCount - 3 do
      begin
        if StrToInt(sgStatistic.Cells[1, iPos2]) <
          StrToInt(sgStatistic.Cells[1, iPos1]) then
        begin
          slRow.Clear;
          for iCol := 0 to sgStatistic.ColCount - 1 do
          begin
            slRow.Add(sgStatistic.Cells[iCol, iPos1]);
          end;
          sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
          for iCol := 0 to slRow.Count - 1 do
          begin
            sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
          end;
        end
        else
        if StrToInt(sgStatistic.Cells[1, iPos2]) =
          StrToInt(sgStatistic.Cells[1, iPos1]) then
        begin
          if cbCollatedSort.Checked = True then
          begin
            if UTF8CompareStrCollated(sgStatistic.Cells[0, iPos2],
              sgStatistic.Cells[0, iPos1]) > 0 then
            begin
              slRow.Clear;
              for iCol := 0 to sgStatistic.ColCount - 1 do
              begin
                slRow.Add(sgStatistic.Cells[iCol, iPos1]);
              end;
              sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
              for iCol := 0 to slRow.Count - 1 do
              begin
                sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
              end;
            end;
          end
          else
          begin
            if UTF8CompareStr(sgStatistic.Cells[0, iPos2],
              sgStatistic.Cells[0, iPos1]) > 0 then
            begin
              slRow.Clear;
              for iCol := 0 to sgStatistic.ColCount - 1 do
              begin
                slRow.Add(sgStatistic.Cells[iCol, iPos1]);
              end;
              sgStatistic.Rows[iPos1] := sgStatistic.Rows[iPos2];
              for iCol := 0 to slRow.Count - 1 do
              begin
                sgStatistic.Cells[iCol, iPos2] := slRow[iCol];
              end;
            end;
          end;
        end;
      end;
    end;
  finally
    slRow.Free;
  end;
end;

procedure TfmMain.miStatisticSaveClick(Sender: TObject);
begin
  // Salve statistic
  pcMain.ActivePage := tsStatistic;
  if sgStatistic.RowCount = 0 then
  begin
    MessageDlg(msg028,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  sdSaveDialog.Title := dgc011;
  sdSaveDialog.Filter := dgc012;
  sdSaveDialog.DefaultExt := '.csv';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      sgStatistic.SaveToCSVFile(sdSaveDialog.FileName);
      sbStatusBar.SimpleText := sbr006;
    except
      MessageDlg(msg029,
        mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.miDiagramAllWordNoBookClick(Sender: TObject);
begin
  // Create diagram of all words no bookmarks
  pcMain.ActivePage := tsDiagram;
  CreateDiagramAllWordsNoBook;
end;

procedure TfmMain.miDiagramAllWordsBookClick(Sender: TObject);
begin
  // Create diagram of all words in bookmarks
  pcMain.ActivePage := tsDiagram;
  CreateDiagramAllWordsBook;
end;

procedure TfmMain.miDiagramSingleWordsBookClick(Sender: TObject);
begin
  // Create diagram of single words in bookmarks
  pcMain.ActivePage := tsDiagram;
  CreateDiagramSingleWordsBook;
end;

procedure TfmMain.miDiagramSaveClick(Sender: TObject);
begin
  // Save diagram
  pcMain.ActivePage := tsDiagram;
  if chChart.Visible = False then
  begin
    MessageDlg(msg030, mtWarning, [mbOK], 0);
    Abort;
  end;
  sdSaveDialog.Title := dgc013;
  sdSaveDialog.Filter := dgc014;
  sdSaveDialog.DefaultExt := '.jpg';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
    try
      try
        chChart.Height := 1000;
        if UTF8LowerCase(ExtractFileExt(sdSaveDialog.FileName)) = '.png' then
        begin
          ChChart.SaveToFile(TPortableNetworkGraphic, sdSaveDialog.FileName);
        end
        else
        begin
          chChart.SaveToFile(TJpegImage, sdSaveDialog.FileName);
        end;
        sbStatusBar.SimpleText := sbr007;
      except
        MessageDlg(msg031,
          mtWarning, [mbOK], 0);
      end;
    finally
      chChart.Height := sbDiagram.Height - 10;
    end;
end;

procedure TfmMain.miDiagramShowValClick(Sender: TObject);
begin
  // Show values
  chChartBarSeries1.Marks.Visible := miDiagramShowVal.Checked;
  chChartBarSeries2.Marks.Visible := miDiagramShowVal.Checked;
  chChartBarSeries3.Marks.Visible := miDiagramShowVal.Checked;
  chChartBarSeries4.Marks.Visible := miDiagramShowVal.Checked;
  chChartBarSeries5.Marks.Visible := miDiagramShowVal.Checked;
  if ((miDiagramShowGrid.Checked = False) and (miDiagramShowVal.Checked = False)) then
  begin
    miDiagramShowGrid.Checked := True;
    miDiagramShowGridClick(nil);
  end;
end;

procedure TfmMain.miDiagramShowGridClick(Sender: TObject);
begin
  // Show axis
  chChart.AxisList[0].Visible := miDiagramShowGrid.Checked;
  chChart.AxisList[1].Grid.Visible := miDiagramShowGrid.Checked;
  if ((miDiagramShowGrid.Checked = False) and (miDiagramShowVal.Checked = False)) then
  begin
    miDiagramShowVal.Checked := True;
    miDiagramShowValClick(nil);
  end;
end;

procedure TfmMain.miDiaZoomInClick(Sender: TObject);
begin
  // Zoom in
  if chChart.Width < 9899 then
  begin
    chChart.Width := chChart.Width + 100;
  end;
end;

procedure TfmMain.miDiaZoomOutClick(Sender: TObject);
begin
  // Zoom out
  if chChart.Width > sbDiagram.Width + 97 then
  begin
    chChart.Width := chChart.Width - 100;
  end
  else
  begin
    chChart.Width := sbDiagram.Width - 3;
  end;
end;

procedure TfmMain.miDiaZoomNormalClick(Sender: TObject);
begin
  // Normal zoom
  chChart.Width := sbDiagram.Width - 3;
end;

procedure TfmMain.miManualClick(Sender: TObject);
begin
  // Show manual
  if OpenDocument(ExtractFileDir(Application.ExeName) + DirectorySeparator +
    man001) = False then
  begin
    MessageDlg(msg032 + LineEnding + 'https://sites.google.com/site/wordstatix.',
      mtInformation, [mbOK], 0);
  end;
end;

procedure TfmMain.miCopyrightFormClick(Sender: TObject);
begin
  // Copyright
  if stCpr1 <> '' then
  begin
    fmCopyright.lbCopyrightSubTitle.Caption := stCpr1;
  end;
  if stCpr2 <> '' then
  begin
    fmCopyright.lbCopyrightAuthor1.Caption := stCpr2;
  end;
  if stCpr3 <> '' then
  begin
    fmCopyright.lbSite.Caption := stCpr3;
  end;
  fmCopyright.ShowModal;
end;

// ********************************************************** //
// ******************* General procedures ******************* //
// ********************************************************** //

procedure TfmMain.CreateBookmarks(blSetCursor: boolean);
var
  iStart, iEnd, i, n: integer;
  stText, stOldBookmark: string;
  blCommasSpaces: boolean;
  slBookmark: TStringList;
begin
  // Set boomarks in the text
  stText := meText.Text;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    slBookmark := TStringList.Create;
    while UTF8Pos('[[', stText) > 0 do
    begin
      iStart := UTF8Pos('[[', stText) + 2;
      iEnd := UTF8Pos(']]', stText);
      if iEnd < iStart then
      begin
        lbBookmarks.Clear;
        if blSetCursor = True then
        begin
          Screen.Cursor := crDefault;
        end;
        MessageDlg(msg033, mtWarning, [mbOK], 0);
        Exit;
      end
      else if iStart + 50 < iEnd then
      begin
        lbBookmarks.Clear;
        if blSetCursor = True then
        begin
          Screen.Cursor := crDefault;
        end;
        MessageDlg(msg034, mtWarning, [mbOK], 0);
        Exit;
      end
      else
      if UTF8Pos(LineEnding, UTF8Copy(stText, iStart, iEnd - iStart)) > 0 then
      begin
        lbBookmarks.Clear;
        if blSetCursor = True then
        begin
          Screen.Cursor := crDefault;
        end;
        MessageDlg(msg035, mtWarning, [mbOK], 0);
        Exit;
      end
      else if UTF8Copy(stText, iStart, iEnd - iStart) <> '' then
      begin
        slBookmark.Add(UTF8Copy(stText, iStart, iEnd - iStart));
      end;
      stText := UTF8Copy(stText, UTF8Pos(']]', stText) + 2, UTF8Length(stText));
    end;
    blCommasSpaces := False;
    for i := 0 to slBookmark.Count - 1 do
    begin
      if slBookmark[i] = '*' then
      begin
        slBookmark[i] := '.';
        meText.Text := StringReplace(meText.Text, '[[*]]', '[[.]]', []);
        blCommasSpaces := True;
      end;
      if ((UTF8Pos(',', slBookmark[i]) > 0) or
        (UTF8Pos(' ', slBookmark[i]) > 0)) then
      begin
        stOldBookmark := slBookmark[i];
        slBookmark[i] :=
          StringReplace(slBookmark[i], ',', '', [rfReplaceAll]);
        slBookmark[i] :=
          StringReplace(slBookmark[i], ' ', '_', [rfReplaceAll]);
        meText.Text := StringReplace(meText.Text, '[[' + stOldBookmark +
          ']]', '[[' + slBookmark[i] + ']]', [rfIgnoreCase]);
        blCommasSpaces := True;
      end;
    end;
    if slBookmark.Count > 1 then
    begin
      for i := 0 to slBookmark.Count - 2 do
      begin
        for n := i + 1 to slBookmark.Count - 1 do
        begin
          if UTF8LowerCase(slBookmark[i]) = UTF8LowerCase(slBookmark[n]) then
          begin
            lbBookmarks.Clear;
            if blSetCursor = True then
            begin
              Screen.Cursor := crDefault;
            end;
            MessageDlg(msg036 + ' "' + slBookmark[i] + '" ' + msg037,
              mtWarning, [mbOK], 0);
            Exit;
          end;
        end;
      end;
    end;
    lbBookmarks.Clear;
    for i := 0 to slBookmark.Count - 1 do
    begin
      lbBookmarks.Items.Add(slBookmark[i]);
    end;
    if blCommasSpaces = True then
    begin
      if blSetCursor = True then
      begin
        Screen.Cursor := crDefault;
      end;
      MessageDlg(msg038,
        mtWarning, [mbOK], 0);
    end;
  finally
    slBookmark.Free;
    if blSetCursor = True then
    begin
      Screen.Cursor := crDefault;
    end;
  end;
end;

function TfmMain.FindNextWord(blMessage: boolean): boolean;
begin
  // Find next word
  if edFind.Text <> '' then
  begin
    if UTF8Pos(UTF8LowerCase(edFind.Text), UTF8LowerCase(meText.Text),
      meText.SelStart + meText.SelLength + 1) > 0 then
    begin
      Result := True;
      meText.SelStart := UTF8Pos(UTF8LowerCase(edFind.Text),
        UTF8LowerCase(meText.Text), meText.SelStart + meText.SelLength + 1) - 1;
      meText.SelLength := UTF8Length(edFind.Text);
      meText.SetFocus;
    end
    else
    begin
      Result := False;
      if blMessage = True then
      begin
        MessageDlg(msg039,
          mtInformation, [mbOK], 0);
      end;
    end;
  end;
end;

procedure TfmMain.CreateConc(stStartText: string);
var
  slNoWord: TStringList;
  stWord, stBookmark: string;
  slText: TStringList;
  i, iTestNum: integer;
  flNewWord: boolean;
  flTmpFile: TextFile;

  // Create concordance
  procedure EvalStopProcess; inline;
  begin
    if blStopConcordance = True then
    begin
      sgWordList.RowCount := 1;
      lsContext.Clear;
      sgStatistic.RowCount := 0;
      sgStatistic.ColCount := 0;
      SetLength(arWordList, 0);
      SetLength(arWordTextual, 0);
      iWordsUsed := 0;
      iWordsTotal := 0;
      chChart.Visible := False;
      chChartBarSeries1.Active := False;
      chChartBarSeries2.Active := False;
      chChartBarSeries3.Active := False;
      chChartBarSeries4.Active := False;
      chChartBarSeries5.Active := False;
      cbComboDiag1.Clear;
      cbComboDiag2.Clear;
      cbComboDiag3.Clear;
      cbComboDiag4.Clear;
      cbComboDiag5.Clear;
      clDiagBookmark.Clear;
      sbStatusBar.SimpleText := sbr008;
      EnableMenuItems;
      blConInProc := False;
      Screen.Cursor := crDefault;
      Abort;
    end;
  end;

begin
  if meText.Text = '' then
  begin
    pcMain.ActivePage := tsFile;
    MessageDlg(msg040, mtWarning, [mbOK], 0);
    Abort;
  end
  else
  if sgWordList.RowCount > 1 then
  begin
    if MessageDlg(msg041, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
      Abort;
  end;
  ttStart := Now;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  sbStatusBar.SimpleText := sbr009;
  CreateBookmarks(False);
  blStopConcordance := False;
  blConInProc := True;
  DisableMenuItems;
  sgStatistic.RowCount := 0;
  sgStatistic.ColCount := 0;
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  cbComboDiag1.Clear;
  cbComboDiag2.Clear;
  cbComboDiag3.Clear;
  cbComboDiag4.Clear;
  cbComboDiag5.Clear;
  clDiagBookmark.Clear;
  sgWordList.Enabled := False;
  meSkipList.Text := CleanField(meSkipList.Text);
  edFltStart.Text := CleanField(edFltStart.Text);
  edFltEnd.Text := CleanField(edFltEnd.Text);
  SetLength(arWordList, 0);
  SetLength(arWordTextual, 0);
  iWordsUsed := 0;
  iWordsTotal := 0;
  stBookmark := '*';
  sgWordList.RowCount := 1;
  lsContext.Clear;
  stStartText := StringReplace(stStartText, #9, ' ', [rfReplaceAll]);
  stStartText := StringReplace(stStartText, #39, #39 + ' ', [rfReplaceAll]);
  stStartText := StringReplace(stStartText, #96, #39 + ' ', [rfReplaceAll]);
  stStartText := StringReplace(stStartText, '’', #39 + ' ', [rfReplaceAll]);
  stStartText := StringReplace(stStartText, '[[', ' [[', [rfReplaceAll]);
  stStartText := StringReplace(stStartText, ']]', ']] ', [rfReplaceAll]);
  while UTF8Pos(LineEnding + ' [[', stStartText) > 0 do
  begin
    stStartText := StringReplace(stStartText,
      LineEnding + ' [[', LineEnding + '[[', [rfReplaceAll]);
  end;
  while UTF8Pos('  ', stStartText) > 0 do
  begin
    stStartText := StringReplace(stStartText, '  ', ' ', [rfReplaceAll]);
  end;
  EvalStopProcess;
  try
    slText := TStringList.Create;
    slText.Delimiter := ' ';
    slText.QuoteChar := ' ';
    slText.DelimitedText := stStartText;
    sltext.SaveToFile(GetTempDir + DirectorySeparator + 'wrdstxtmp');
    iWordsStartTot := slText.Count;
    stStartText := '';
  finally
    slText.Free;
  end;
  try
    SetLength(arWordTextual, iWordsStartTot);
    iWordsStartTot := iWordsStartTot - lbBookmarks.Count;
    slNoWord := TStringList.Create;
    slNoWord.CaseSensitive := False;
    slNoWord.CommaText := meSkipList.Text;
    sbStatusBar.SimpleText :=
      sbr010 + ' ' + FormatFloat('#,##0', iWordsTotal) + ' ' +
      sbr011 + ' ' + FormatFloat('#,##0', iWordsStartTot) + ' - ' +
      sbr012 + ' ' + FormatFloat('#,##0', iWordsUsed) + '.';
    AssignFile(flTmpFile, GetTempDir + DirectorySeparator + 'wrdstxtmp');
    Reset(flTmpFile);
    while not EOF(flTmpFile) do
    begin
      ReadLn(flTmpFile, stWord);
      if UTF8Pos('[[', stWord) > 0 then
      begin
        stBookmark := UTF8Copy(stWord, UTF8Pos('[[', stWord) + 2,
          UTF8Pos(']]', stWord) - UTF8Pos('[[', stWord) - 2);
        if UTF8Length(stBookmark) > 20 then
          stBookmark := UTF8Copy(stBookmark, 1, 20) + '...';
        if stBookmark = '' then
        begin
          stBookmark := '?';
        end;
        Application.ProcessMessages;
        Continue;
      end;
      arWordTextual[iWordsTotal].stRecWordTextual := stWord;
      Inc(iWordsTotal);
      for i := 33 to 38 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      for i := 40 to 47 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      for i := 58 to 64 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      for i := 91 to 95 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      for i := 123 to 127 do
      begin
        stWord := StringReplace(stWord, Chr(i), '', [rfReplaceAll]);
      end;
      stWord := StringReplace(stWord, '«', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '»', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '“', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '”', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '–', '', [rfReplaceAll]);
      stWord := StringReplace(stWord, '  ', ' ', [rfReplaceAll]);
      if cbSkipNumbers.Checked = True then
      begin
        if TryStrToInt(stWord, iTestNum) = True then
        begin
          Application.ProcessMessages;
          Continue;
        end;
      end;
      if slNoWord.IndexOf(stWord) < 0 then
      begin
        if ((stWord <> ' ') and (stWord <> '') and (stWord <> #39) and
          (CheckFilter(stWord) = True)) then
        begin
          flNewWord := True;
          if Length(arWordList) > 0 then
          begin
            for i := 0 to Length(arWordList) - 1 do
            begin
              if UTF8LowerCase(stWord) = arWordList[i].stRecWord then
              begin
                flNewWord := False;
                Break;
              end;
            end;
          end;
          if flNewWord = True then
          begin
            SetLength(arWordList, Length(arWordList) + 1);
            arWordList[Length(arWordList) - 1].stRecWord := UTF8LowerCase(stWord);
            arWordList[Length(arWordList) - 1].iRecFreq := 1;
            arWordList[Length(arWordList) - 1].stRecWordPos :=
              arWordList[Length(arWordList) - 1].stRecWordPos +
              IntToStr(iWordsTotal - 1) + ',';
            if stBookmark <> '' then
            begin
              arWordList[Length(arWordList) - 1].stRecBookPos :=
                arWordList[Length(arWordList) - 1].stRecBookPos +
                stBookmark + ',';
            end;
            Inc(iWordsUsed);
          end
          else
          begin
            arWordList[i].iRecFreq := arWordList[i].iRecFreq + 1;
            arWordList[i].stRecWordPos :=
              arWordList[i].stRecWordPos + IntToStr(iWordsTotal - 1) + ',';
            if stBookmark <> '' then
            begin
              arWordList[i].stRecBookPos :=
                arWordList[i].stRecBookPos + stBookmark + ',';
            end;
          end;
        end;
      end;
      if iWordsStartTot > 10000 then
      begin
        if iWordsTotal mod 5000 = 0 then
        begin
          sbStatusBar.SimpleText :=
            sbr013 + ' ' + FormatFloat('#,##0', iWordsTotal) + ' ' +
            sbr014 + ' ' + FormatFloat('#,##0', iWordsStartTot) + ' - ' +
            sbr015 + ' ' + FormatFloat('#,##0', iWordsUsed) + '.';
        end;
      end;
      Application.ProcessMessages;
      EvalStopProcess;
    end;
    CompileGrid(True);
  finally
    CloseFile(flTmpFile);
    slNoWord.Free;
    sgWordList.Enabled := True;
    EnableMenuItems;
    blConInProc := False;
    Screen.Cursor := crDefault;
  end;
  if FileExistsUTF8(GetTempDir + DirectorySeparator + 'wrdstxtmp') then
  begin
    DeleteFileUTF8(GetTempDir + DirectorySeparator + 'wrdstxtmp');
  end;
  blTextModConc := False;
  ttEnd := Now;
  sbStatusBar.SimpleText :=
    sbr016 + ' ' + FormatFloat('#,##0', iWordsTotal) + ' - ' + sbr017 +
    ' ' + FormatFloat('#,##0', iWordsUsed) + '. ' + sbr018 + ' ' +
    FormatDateTime('hh:nn:ss', ttEnd - ttStart) + '.';
  MessageDlg(msg042 + LineEnding + LineEnding + msg043 + ' ' +
    FormatFloat('#,##0', iWordsTotal) + '.' + LineEnding + msg044 +
    ' ' + FormatFloat('#,##0', iWordsUsed) + '.' + LineEnding + msg045 +
    ' ' + FormatDateTime('hh:nn:ss', ttEnd - ttStart) + '.',
    mtInformation, [mbOK], 0);
end;

function TfmMain.CreateContext(GridRow: integer; blSetCursor: boolean): string;
var
  stPos, stBookList, stBookmark, stContext, stDotsBf, stDotsAf: string;
  iPos, i: integer;
begin
  // Create context
  if blSetCursor = True then
  begin
    Screen.Cursor := crHourGlass;
  end;
  Application.ProcessMessages;
  Result := '';
  stPos := sgWordList.Cells[3, GridRow];
  stBookList := sgWordList.Cells[4, GridRow];
  stBookmark := '';
  try
    while Pos(',', stPos) > 0 do
    begin
      iPos := StrToInt(Copy(stPos, 1, Pos(',', stPos) - 1));
      stBookmark := Copy(stBookList, 1, Pos(',', stBookList) - 1);
      stContext := '';
      stDotsBf := '... ';
      stDotsAf := '...';
      for i := iPos - StrToInt(edWordsCont.Text)
        to iPos + StrToInt(edWordsCont.Text) do
      begin
        if i < 0 then
        begin
          stDotsBf := '';
          Continue;
        end;
        if i > Length(arWordTextual) - 1 then
        begin
          stDotsAf := '';
          Continue;
        end;
        if ((i = iPos - 1) or (i = iPos)) then
          stContext := stContext + arWordTextual[i].stRecWordTextual + ' • '
        else
          stContext := stContext + arWordTextual[i].stRecWordTextual + ' ';
      end;
      stContext := StringReplace(stContext, #39 + ' ', #39, [rfReplaceAll]);
      Result := Result + LineEnding + '[' + stBookmark + '] ' +
        stDotsBf + stContext + stDotsAf;
      stPos := Copy(stPos, Pos(',', stPos) + 1, Length(stPos));
      stBookList := Copy(stBookList, Pos(',', stBookList) + 1, Length(stBookList));
    end;
    Result := UTF8Copy(Result, 2, UTF8Length(Result));
  finally
    if blSetCursor = True then
    begin
      Screen.Cursor := crDefault;
    end;
  end;
end;

procedure TfmMain.AddSkipWord;
var
  WordsList: TStringList;
begin
  // Add selected word to skip list
  if sgWordList.RowCount > 1 then
  begin
    if sgWordList.RowHeights[sgWordList.Row] > 0 then
    begin
      if sgWordList.Cells[0, sgWordList.Row] <> '' then
      begin
        meSkipList.Text := CleanField(meSkipList.Text);
        try
          WordsList := TStringList.Create;
          WordsList.CommaText := meSkipList.Text;
          if WordsList.IndexOf(sgWordList.Cells[0, sgWordList.Row]) < 0 then
          begin
            meSkipList.Text :=
              meSkipList.Text + ', ' + sgWordList.Cells[0, sgWordList.Row];
            meSkipList.Text := CleanField(meSkipList.Text);
            sgWordList.DeleteRow(sgWordList.Row);
            blSortMod := True;
          end;
        finally
          WordsList.Free;
        end;
      end;
    end;
    lsContext.Clear;
    if sgWordList.RowCount > 1 then
    begin
      if sgWordList.RowHeights[sgWordList.Row] > 0 then
      begin
        lsContext.Items.Text := CreateContext(sgWordList.Row, True);
        {$ifdef Windows}
        // Due to a bug in Windows...
        if lsContext.Items[0] = '' then
          lsContext.Items.Delete(0);
        {$endif}
        lsContext.ItemIndex := 0;
      end;
    end;
  end
  else
  begin
    MessageDlg(msg046, mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.CompileGrid(blSort: boolean);
var
  i: integer;
begin
  // Compile grid of words
  if Length(arWordList) = 0 then
  begin
    sgWordList.RowCount := 1;
    lsContext.Clear;
    sbStatusBar.SimpleText := sbr019;
  end
  else
  begin
    try
      Screen.Cursor := crHourGlass;
      blStopConcordance := False;
      sgWordList.Enabled := False;
      miConcodanceShowSelected.Checked := False;
      if blSort = True then
      begin
        sbStatusBar.SimpleText := sbr020;
        Application.ProcessMessages;
        SortWordFreq(arWordList, rgSortBy.ItemIndex);
      end;
      sbStatusBar.SimpleText := sbr021;
      Application.ProcessMessages;
      sgWordList.RowCount := 1;
      sgWordList.RowCount := Length(arWordList) + 1;
      for i := 0 to Length(arWordList) - 1 do
      begin
        sgWordList.Cells[0, i + 1] := arWordList[i].stRecWord;
        sgWordList.Cells[1, i + 1] := IntToStr(arWordList[i].iRecFreq);
        sgWordList.Cells[2, i + 1] := '0';
        sgWordList.Cells[3, i + 1] := arWordList[i].stRecWordPos;
        sgWordList.Cells[4, i + 1] := arWordList[i].stRecBookPos;
        Application.ProcessMessages;
      end;
      lsContext.Clear;
      lsContext.Items.Text := CreateContext(sgWordList.Row, True);
    {$ifdef Windows}
      // Due to a bug in Windows...
      if lsContext.Items[0] = '' then
        lsContext.Items.Delete(0);
    {$endif}
      lsContext.ItemIndex := 0;
      sbStatusBar.SimpleText :=
        sbr022 + ' ' + FormatFloat('#,##0', iWordsTotal) + ' - ' +
        sbr023 + ' ' + FormatFloat('#,##0', iWordsUsed) + '.';
    finally
      sgWordList.Enabled := True;
      Screen.Cursor := crDefault;
    end;
  end;
  blSortMod := False;
end;

function TfmMain.CheckFilter(stWord: string): boolean;
var
  slFltStart, slFltEnd: TStringList;
  i: integer;
  blStart, blEnd: boolean;
begin
  // Check the filtr on words
  Result := False;
  if ((edFltStart.Text = '') and (edFltEnd.Text = '')) then
  begin
    Result := True;
  end
  else
  begin
    blStart := False;
    blEnd := False;
    if edFltStart.Text <> '' then
    begin
      try
        slFltStart := TStringList.Create;
        slFltStart.CommaText := edFltStart.Text;
        for i := 0 to slFltStart.Count - 1 do
        begin
          if UTF8LowerCase(UTF8Copy(stWord, 1, UTF8Length(slFltStart[i]))) =
            UTF8LowerCase(slFltStart[i]) then
          begin
            blStart := True;
            Break;
          end;
        end;
      finally
        slFltStart.Free;
      end;
    end;
    if edFltEnd.Text <> '' then
    begin
      try
        slFltEnd := TStringList.Create;
        slFltEnd.CommaText := edFltEnd.Text;
        for i := 0 to slFltEnd.Count - 1 do
        begin
          if UTF8LowerCase(UTF8Copy(stWord, UTF8Length(stWord) -
            UTF8Length(slFltEnd[i]) + 1, UTF8Length(stWord))) =
            UTF8LowerCase(slFltEnd[i]) then
          begin
            blEnd := True;
            Break;
          end;
        end;
      finally
        slFltEnd.Free;
      end;
    end;
    if rgCondFlt.ItemIndex = 0 then
    begin
      if ((blStart = True) or (blEnd = True)) then
      begin
        Result := True;
      end
      else
      begin
        Result := False;
      end;
    end
    else
    begin
      if edFltStart.Text = '' then
      begin
        blStart := True;
      end;
      if edFltEnd.Text = '' then
      begin
        blEnd := True;
      end;
      if ((blStart = True) and (blEnd = True)) then
      begin
        Result := True;
      end
      else
      begin
        Result := False;
      end;
    end;
  end;
end;

procedure TfmMain.SortWordFreq(var arWordList: array of TRecWordList; flField: shortint);

  procedure EvalStopSort; inline;
  begin
    // Stop sort words
    if blStopConcordance = True then
    begin
      sgWordList.RowCount := 1;
      lsContext.Clear;
      sgStatistic.RowCount := 0;
      sgStatistic.ColCount := 0;
      chChart.Visible := False;
      chChartBarSeries1.Active := False;
      chChartBarSeries2.Active := False;
      chChartBarSeries3.Active := False;
      chChartBarSeries4.Active := False;
      chChartBarSeries5.Active := False;
      cbComboDiag1.Clear;
      cbComboDiag2.Clear;
      cbComboDiag3.Clear;
      cbComboDiag4.Clear;
      cbComboDiag5.Clear;
      clDiagBookmark.Clear;
      sbStatusBar.SimpleText := sbr024;
      EnableMenuItems;
      blConInProc := False;
      blSortMod := False;
      Screen.Cursor := crDefault;
      Abort;
    end;
  end;

var
  iPos1, iPos2: integer;
  rcTempRec: TRecWordList;
begin
  // Sort words
  for iPos1 := 0 to Length(arWordList) - 1 do
  begin
    for iPos2 := 0 to Length(arWordList) - 2 do
    begin
      if flField = 0 then
      begin
        if arWordList[iPos2].iRecFreq < arWordList[iPos2 + 1].iRecFreq then
        begin
          rcTempRec := arWordList[iPos2];
          arWordList[iPos2] := arWordList[iPos2 + 1];
          arWordList[iPos2 + 1] := rcTempRec;
        end
        else
        begin
          if arWordList[iPos2].iRecFreq = arWordList[iPos2 + 1].iRecFreq then
          begin
            if cbCollatedSort.Checked = True then
            begin
              if UTF8CompareStrCollated(arWordList[iPos2].stRecWord,
                arWordList[iPos2 + 1].stRecWord) > 0 then
              begin
                rcTempRec := arWordList[iPos2];
                arWordList[iPos2] := arWordList[iPos2 + 1];
                arWordList[iPos2 + 1] := rcTempRec;
              end;
            end
            else
            begin
              if UTF8CompareStr(arWordList[iPos2].stRecWord,
                arWordList[iPos2 + 1].stRecWord) > 0 then
              begin
                rcTempRec := arWordList[iPos2];
                arWordList[iPos2] := arWordList[iPos2 + 1];
                arWordList[iPos2 + 1] := rcTempRec;
              end;
            end;
          end;
        end;
      end
      else
      begin
        if cbCollatedSort.Checked = True then
        begin
          if UTF8CompareStrCollated(arWordList[iPos2].stRecWord,
            arWordList[iPos2 + 1].stRecWord) > 0 then
          begin
            rcTempRec := arWordList[iPos2];
            arWordList[iPos2] := arWordList[iPos2 + 1];
            arWordList[iPos2 + 1] := rcTempRec;
          end;
        end
        else
        begin
          if UTF8CompareStr(arWordList[iPos2].stRecWord,
            arWordList[iPos2 + 1].stRecWord) > 0 then
          begin
            rcTempRec := arWordList[iPos2];
            arWordList[iPos2] := arWordList[iPos2 + 1];
            arWordList[iPos2 + 1] := rcTempRec;
          end;
        end;
      end;
    end;
    Application.ProcessMessages;
    EvalStopSort;
  end;
end;

function TfmMain.CleanField(myField: string): string;
var
  myText: string;
  i: integer;
  WordsList, WordCurr: TStringList;
begin
  // Clean fields with commas by wrong inputs
  if myField <> '' then
  begin
    myText := myField;
    while ((UTF8Copy(myText, UTF8Length(myText), 1) = ',') or
        (UTF8Copy(myText, UTF8Length(myText), 1) = ' ')) do
    begin
      myText := UTF8Copy(myText, 1, UTF8Length(myText) - 1);
    end;
    myText := StringReplace(myText, ',', ', ', [rfReplaceAll]);
    myText := StringReplace(myText, ' ,', ',', [rfReplaceAll]);
    while Pos('  ', myText) > 0 do
    begin
      myText := StringReplace(myText, '  ', ' ', [rfReplaceAll]);
    end;
    while UTF8Copy(myText, 1, 1) = ' ' do
    begin
      myText := UTF8Copy(myText, 2, UTF8Length(myText));
    end;
    Result := myText;
    if myText <> '' then
    begin
      try
        WordsList := TStringList.Create;
        WordCurr := TStringList.Create;
        WordCurr.Text := StringReplace(myText, ', ', LineEnding, [rfReplaceAll]);
        WordCurr.Sort;
        for i := 0 to WordCurr.Count - 1 do
        begin
          if WordsList.IndexOf(WordCurr[i]) < 0 then
          begin
            if WordCurr[i] <> '' then
              WordsList.Add(WordCurr[i]);
          end;
        end;
        myText := WordsList.Text;
        Result := StringReplace(myText, LineEnding, ', ', [rfReplaceAll]);
        Result := UTF8Copy(Result, 1, UTF8Length(Result) - 2);
      finally
        WordsList.Free;
        WordCurr.Free;
      end;
    end;
  end
  else
  begin
    Result := '';
  end;
end;

procedure TfmMain.SaveConcordance;
var
  flConc: TextFile;
  i: integer;
  stConditions: string;
begin
  // Save concordance
  if meText.Text = '' then
  begin
    MessageDlg(msg047, mtWarning, [mbOK], 0);
    Abort;
  end;
  if sgWordList.RowCount = 1 then
  begin
    MessageDlg(msg048,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if blTextModConc = True then
  begin
    MessageDlg(msg049,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  sdSaveDialog.Title := 'Save concordance';
  sdSaveDialog.Filter := 'WordStatix files|*.wsx|All files|*';
  sdSaveDialog.DefaultExt := '.wsx';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
  begin
    try
      try
        Screen.Cursor := crHourGlass;
        Application.ProcessMessages;
        AssignFile(flConc, sdSaveDialog.FileName);
        Rewrite(flConc);
        WriteLn(flConc, '[WordStatixText]');
        for i := 0 to meText.Lines.Count - 1 do
        begin
          WriteLn(flConc, meText.Lines[i]);
        end;
        WriteLn(flConc, '[WordStatixList]');
        for i := 0 to Length(arWordList) - 1 do
        begin
          WriteLn(flConc, arWordList[i].stRecWord + #9 +
            IntToStr(arWordList[i].iRecFreq) + #9 + arWordList[i].stRecWordPos +
            #9 + arWordList[i].stRecBookPos);
        end;
        WriteLn(flConc, '[WordStatixSeries]');
        for i := 0 to Length(arWordTextual) - 1 do
        begin
          WriteLn(flConc, arWordTextual[i].stRecWordTextual);
        end;
        WriteLn(flConc, '[WordStatixConditions]');
        stConditions := '';
        stConditions := IntToStr(rgSortBy.ItemIndex) + #9;
        if cbCollatedSort.Checked = True then
        begin
          stConditions := stConditions + 'CST' + #9;
        end
        else
        begin
          stConditions := stConditions + 'CSF' + #9;
        end;
        if cbSkipNumbers.Checked = True then
        begin
          stConditions := stConditions + 'SNT' + #9;
        end
        else
        begin
          stConditions := stConditions + 'SNF' + #9;
        end;
        stConditions := stConditions + meSkipList.Lines[0] + #9;
        stConditions := stConditions + edFltStart.Text + #9;
        stConditions := stConditions + edFltEnd.Text + #9;
        stConditions := stConditions + IntToStr(rgCondFlt.ItemIndex) + #9;
        WriteLn(flConc, stCOnditions);
        WriteLn(flConc, '[WordStatixResults]');
        WriteLn(flConc, IntToStr(iWordsTotal) + #9 + IntToStr(iWordsUsed));
        sbStatusBar.SimpleText := sbr025;
      finally
        CloseFile(flConc);
        Screen.Cursor := crDefault;
      end;
    except
      MessageDlg(msg050,
        mtWarning, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.OpenConcordance;
var
  flConc: TextFile;
  stLine: string;
  iSection: smallint;
begin
  // Open concordance
  pcMain.ActivePage := tsFile;
  if ((meText.Text <> '') and (blTextModSave = True)) then
  begin
    if MessageDlg(msg051, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    begin
      Abort;
    end;
  end;
  odOpenDialog.Title := dgc015;
  odOpenDialog.Filter := dgc016;
  odOpenDialog.DefaultExt := '.wsx';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute then
  begin
    try
      Screen.Cursor := crHourGlass;
      sbStatusBar.SimpleText := sbr026;
      pcMain.ActivePage := tsConcordance;
      Application.ProcessMessages;
      ttStart := Now;
      SetLength(arWordList, 0);
      SetLength(arWordTextual, 0);
      iWordsUsed := 0;
      iWordsTotal := 0;
      meText.Clear;
      lbBookmarks.Clear;
      edFind.Clear;
      edReplace.Clear;
      sgWordList.RowCount := 1;
      lsContext.Clear;
      edFltStart.Clear;
      edFltEnd.Clear;
      sgStatistic.RowCount := 0;
      sgStatistic.ColCount := 0;
      chChart.Visible := False;
      chChartBarSeries1.Active := False;
      chChartBarSeries2.Active := False;
      chChartBarSeries3.Active := False;
      chChartBarSeries4.Active := False;
      chChartBarSeries5.Active := False;
      cbComboDiag1.Clear;
      cbComboDiag2.Clear;
      cbComboDiag3.Clear;
      cbComboDiag4.Clear;
      cbComboDiag5.Clear;
      clDiagBookmark.Clear;
      iSection := -1;
      AssignFile(flConc, odOpenDialog.FileName);
      Reset(flConc);
      try
        while not EOF(flConc) do
        begin
          ReadLn(flConc, stLine);
          if stLine = '[WordStatixText]' then
          begin
            iSection := 0;
          end
          else if stLine = '[WordStatixList]' then
          begin
            iSection := 1;
          end
          else if stLine = '[WordStatixSeries]' then
          begin
            iSection := 2;
          end
          else if stLine = '[WordStatixConditions]' then
          begin
            iSection := 3;
          end
          else if stLine = '[WordStatixResults]' then
          begin
            iSection := 4;
          end
          else if iSection = 0 then
          begin
            meText.Lines.Add(stLine);
          end
          else if iSection = 1 then
          begin
            SetLength(arWordList, Length(arWordList) + 1);
            arWordList[Length(arWordList) - 1].stRecWord :=
              UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            arWordList[Length(arWordList) - 1].iRecFreq :=
              StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            arWordList[Length(arWordList) - 1].stRecWordPos :=
              UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            arWordList[Length(arWordList) - 1].stRecBookPos :=
              UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
          end
          else if iSection = 2 then
          begin
            SetLength(arWordTextual, Length(arWordTextual) + 1);
            arWordTextual[Length(arWordTextual) - 1].stRecWordTextual := stLine;
          end
          else if iSection = 3 then
          begin
            rgSortBy.ItemIndex :=
              StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            if UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1) = 'CST' then
            begin
              cbCollatedSort.Checked := True;
            end
            else if UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1) = 'CSF' then
            begin
              cbCollatedSort.Checked := False;
            end;
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            if UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1) = 'SNT' then
            begin
              cbSkipNumbers.Checked := True;
            end
            else
            if UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1) = 'SNF' then
            begin
              cbSkipNumbers.Checked := False;
            end;
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            meSkipList.Text := UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            edFltStart.Text := UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            edFltEnd.Text := UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1);
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            rgCondFlt.ItemIndex :=
              StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
          end
          else if iSection = 4 then
          begin
            iWordsTotal := StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
            stLine := UTF8Copy(stLine, UTF8Pos(#9, stLine) + 1,
              UTF8Length(stLine));
            iWordsUsed := StrToInt(UTF8Copy(stLine, 1, UTF8Pos(#9, stLine) - 1));
          end
          else
          begin
            MessageDlg(msg052,
              mtWarning, [mbOK], 0);
            Exit;
          end;
        end;
      finally
        CloseFile(flConc);
        Screen.Cursor := crDefault;
      end;
      CreateBookmarks(False);
      CompileGrid(False);
      blTextModSave := False;
      blTextModConc := False;
      miFileSave.Enabled := blTextModSave;
      stFileName := ExtractFileNameWithoutExt(odOpenDialog.FileName) + '.txt';
      fmMain.Caption := 'WordStatix - ' + stFileName;
      ttEnd := Now;
      sbStatusBar.SimpleText :=
        sbr027 + ' ' + FormatFloat('#,##0', iWordsTotal) + ' - ' +
        sbr028 + ' ' + FormatFloat('#,##0', iWordsUsed) + '. ' +
        sbr029 + ' ' + FormatDateTime('hh:nn:ss', ttEnd - ttStart) + '.';
    except
      MessageDlg(msg053,
        mtWarning, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.SaveReport(stFileName: string);
var
  ConcFile: TextFile;
  i: integer;
  stTitle: string;
begin
  // Export concordance to report
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    fmInput.Caption := dgc017;
    fmInput.lbTitle.Caption := dgc018;
    if fmInput.ShowModal = mrOK then
    begin
      stTitle := fmInput.edTitle.Text;
    end;
    if stTitle = '' then
      stTitle := dgc019;
    AssignFile(ConcFile, stFileName);
    Rewrite(ConcFile);
    if UTF8LowerCase(ExtractFileExt(stFileName)) = '.html' then
    begin
      WriteLn(ConcFile, '<HTML>');
      WriteLn(ConcFile, '<HEAD>');
      WriteLn(ConcFile, '<meta http-equiv="Content-Type" ' +
        'content="text/html; charset=UTF-8">');
      WriteLn(ConcFile, '<TITLE>' + stTitle + '</TITLE>');
      WriteLn(ConcFile, ' </HEAD>');
      WriteLn(ConcFile, '<BODY>');
      WriteLn(ConcFile, '<H1>' + stTitle + '</H1>');
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        if sgWordList.RowHeights[i] > 0 then
        begin
          WriteLn(ConcFile, '<H2>' + sgWordList.Cells[0, i] + '</H2>');
          WriteLn(ConcFile, '<b>' + sgWordList.Cells[1, i] + ' ' +
            dgc020 + '.</b>');
          WriteLn(ConcFile, '<BR>');
          WriteLn(ConcFile, StringReplace(CreateContext(i, False),
            LineEnding, '<BR>' + LineEnding, [rfReplaceAll]));
          WriteLn(ConcFile, '<BR>');
          WriteLn(ConcFile, '<BR>');
        end;
      end;
      WriteLn(ConcFile, '</BODY>');
      WriteLn(ConcFile, '</HTML>');
    end
    else
    begin
      WriteLn(ConcFile, stTitle);
      WriteLn(ConcFile, '');
      for i := 1 to sgWordList.RowCount - 1 do
      begin
        if sgWordList.RowHeights[i] > 0 then
        begin
          WriteLn(ConcFile, '*** ' + sgWordList.Cells[0, i] + ' ***');
          WriteLn(ConcFile, sgWordList.Cells[1, i] + ' ' + dgc020 + '.');
          WriteLn(ConcFile, '');
          WriteLn(ConcFile, CreateContext(i, False));
          WriteLn(ConcFile, '');
          WriteLn(ConcFile, '');
        end;
      end;
    end;
  finally
    CloseFile(ConcFile);
    Screen.Cursor := crDefault;
  end;
end;

function TfmMain.CleanXML(stXMLText: string): string;
var
  blTag: boolean;
  i: integer;
begin
  // Clean a text from XML/HTML tags
  // Do not use UTF8Lengt and UTF8Copy here
  blTag := False;
  Result := '';
  for i := 1 to Length(stXMLText) do
  begin
    if Copy(stXMLText, i, 1) = '<' then
    begin
      blTag := True;
    end
    else if Copy(stXMLText, i, 1) = '>' then
    begin
      blTag := False;
    end
    else if blTag = False then
    begin
      Result := Result + Copy(stXMLText, i, 1);
    end;
    Application.ProcessMessages;
  end;
  while Copy(Result, 1, 1) = LineEnding do
  begin
    Result := Copy(Result, 2, Length(Result));
  end;
end;

procedure TfmMain.CreateStatistic;
var
  i, n, iWord, iList, iTot: integer;
  slBookmark: TStringList;
begin
  // Create statistic
  if sgWordList.RowCount = 1 then
  begin
    pcMain.ActivePage := tsConcordance;
    MessageDlg(msg054,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    slBookmark := TStringList.Create;
    sgStatistic.RowCount := 1;
    sgStatistic.FixedRows := 1;
    sgStatistic.ColCount := lbBookmarks.Count + 3;
    sgStatistic.FixedCols := 1;
    cbComboDiag1.Clear;
    cbComboDiag2.Clear;
    cbComboDiag3.Clear;
    cbComboDiag4.Clear;
    cbComboDiag5.Clear;
    clDiagBookmark.Clear;
    for i := 1 to sgStatistic.ColCount - 1 do
    begin
      sgStatistic.ColWidths[i] := 70;
    end;
    sgStatistic.Cells[1, 0] := dgc021;
    sgStatistic.Cells[2, 0] := '*';
    clDiagBookmark.Items.Add('*');
    for i := 0 to lbBookmarks.Count - 1 do
    begin
      sgStatistic.Cells[i + 3, 0] := lbBookmarks.Items[i];
      clDiagBookmark.Items.Add(lbBookmarks.Items[i]);
    end;
    clDiagBookmark.CheckAll(cbChecked, False, False);
    for i := 1 to sgWordList.RowCount - 1 do
    begin
      if sgWordList.Cells[2, i] = '1' then
      begin
        sgStatistic.RowCount := sgStatistic.RowCount + 1;
        sgStatistic.Cells[0, sgStatistic.RowCount - 1] :=
          sgWordList.Cells[0, i];
        sgStatistic.Cells[1, sgStatistic.RowCount - 1] :=
          sgWordList.Cells[1, i];
        slBookmark.CommaText := sgWordList.Cells[4, i];
        cbComboDiag1.Items.Add(sgWordList.Cells[0, i]);
        cbComboDiag2.Items.Add(sgWordList.Cells[0, i]);
        cbComboDiag3.Items.Add(sgWordList.Cells[0, i]);
        cbComboDiag4.Items.Add(sgWordList.Cells[0, i]);
        cbComboDiag5.Items.Add(sgWordList.Cells[0, i]);
        for n := 2 to sgStatistic.ColCount - 1 do
        begin
          iWord := 0;
          for iList := 0 to slBookmark.Count - 1 do
          begin
            if UTF8LowerCase(slBookmark[iList]) =
              UTF8LowerCase(sgStatistic.Cells[n, 0]) then
              Inc(iWord);
          end;
          sgStatistic.Cells[n, sgStatistic.RowCount - 1] := IntToStr(iWord);
        end;
      end;
    end;
    cbComboDiag1.ItemIndex := - 1;
    cbComboDiag2.ItemIndex := - 1;
    cbComboDiag3.ItemIndex := - 1;
    cbComboDiag4.ItemIndex := - 1;
    cbComboDiag5.ItemIndex := - 1;
    if sgStatistic.RowCount = 1 then
    begin
      sgStatistic.RowCount := 1;
      sgStatistic.ColCount := 1;
      MessageDlg(msg055,
        mtWarning, [mbOK], 0);
      pcMain.ActivePage := tsConcordance;
      sgWordList.SetFocus;
    end
    else
    begin
      sgStatistic.RowCount := sgStatistic.RowCount + 1;
      sgStatistic.Cells[0, sgStatistic.RowCount - 1] := dgc021;
      for i := 1 to sgStatistic.ColCount - 1 do
      begin
        iTot := 0;
        for n := 1 to sgStatistic.RowCount - 2 do
        begin
          iTot := iTot + StrToInt(sgStatistic.Cells[i, n]);
        end;
        sgStatistic.Cells[i, sgStatistic.RowCount - 1] := IntToStr(iTot);
      end;
    end;
  finally
    Screen.Cursor := crDefault;
    slBookmark.Free;
  end;
end;

procedure TfmMain.CreateDiagramSingleWordsBook;
var
  iRowGridStat, iColGridStat, iNumDiag, iBottomAx, i: integer;
  clColor1, clColor2, clColor3, clColor4, clColor5: TColor;
  slCheckDouble: TStringList;
  blLastField, blSelBookmark: boolean;
begin
  // Create diagram of single words
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  if sgStatistic.RowCount = 0 then
  begin
    pcMain.ActivePage := tsStatistic;
    MessageDlg(msg056,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if clDiagBookmark.Count = 0 then
  begin
    MessageDlg(msg057,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  blSelBookmark := False;
  for i := 0 to clDiagBookmark.Count - 1 do
  begin
    if clDiagBookmark.Checked[i] = True then
    begin
      blSelBookmark := True;
    end;
  end;
  if blSelBookmark = False then
  begin
    MessageDlg(msg058,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if cbComboDiag1.ItemIndex = -1 then
  begin
    MessageDlg(msg059,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  try
    blLastField := False;
    slCheckDouble := TStringList.Create;
    slCheckDouble.Add(cbComboDiag1.Text);
    if cbComboDiag2.ItemIndex > -1 then
    begin
      slCheckDouble.Add(cbComboDiag2.Text);
    end
    else
    begin
      blLastField := True;
    end;
    if cbComboDiag3.ItemIndex > -1 then
    begin
      if blLastField = True then
      begin
        MessageDlg(msg060,
          mtWarning, [mbOK], 0);
        Abort;
      end
      else
      begin
        slCheckDouble.Add(cbComboDiag3.Text);
      end;
    end
    else
    begin
      blLastField := True;
    end;
    if cbComboDiag4.ItemIndex > -1 then
    begin
      if blLastField = True then
      begin
        MessageDlg(msg060,
          mtWarning, [mbOK], 0);
        Abort;
      end
      else
      begin
        slCheckDouble.Add(cbComboDiag4.Text);
      end;
    end
    else
    begin
      blLastField := True;
    end;
    if cbComboDiag5.ItemIndex > -1 then
    begin
      if blLastField = True then
      begin
        MessageDlg(msg060,
          mtWarning, [mbOK], 0);
        Abort;
      end
      else
      begin
        slCheckDouble.Add(cbComboDiag5.Text);
      end;
    end
    else
    begin
      blLastField := True;
    end;
    iNumDiag := slCheckDouble.Count;
    slCheckDouble.Sort;
    for i := 0 to slCheckDouble.Count - 2 do
    begin
      if slCheckDouble[i] = slCheckDouble[i + 1] then
      begin
        MessageDlg(msg061,
          mtWarning, [mbOK], 0);
        Exit;
      end;
    end;
  finally
    slCheckDouble.Free;
  end;
  fmInput.Caption := dgc022;
  fmInput.lbTitle.Caption := dgc023;
  if fmInput.ShowModal = mrOK then
  begin
    stDiagTitle := fmInput.edTitle.Text;
  end;
  if stDiagTitle = '' then
  begin
    chChart.Title.Visible := False;
  end
  else
  begin
    chChart.Title.Text.Clear;
    chChart.Title.Text.Add(stDiagTitle);
    chChart.Title.Visible := True;
  end;
  clColor1 := $8080FF;
  clColor2 := $80FF80;
  clColor3 := $FFB980;
  clColor4 := $80FFFA;
  clColor5 := $E180FF;
  csChartSource1.DataPoints.Clear;
  csChartSource2.DataPoints.Clear;
  csChartSource3.DataPoints.Clear;
  csChartSource4.DataPoints.Clear;
  csChartSource5.DataPoints.Clear;
  chChartBarSeries1.Title := '';
  chChartBarSeries2.Title := '';
  chChartBarSeries3.Title := '';
  chChartBarSeries4.Title := '';
  chChartBarSeries5.Title := '';
  if iNumDiag = 1 then
  begin
    chChartBarSeries1.BarWidthPercent := 70;
    chChartBarSeries1.BarOffsetPercent := 0;
  end
  else if iNumDiag = 2 then
  begin
    chChartBarSeries1.BarWidthPercent := 35;
    chChartBarSeries1.BarOffsetPercent := -20;
    chChartBarSeries2.BarWidthPercent := 35;
    chChartBarSeries2.BarOffsetPercent := 20;
  end
  else if iNumDiag = 3 then
  begin
    chChartBarSeries1.BarWidthPercent := 25;
    chChartBarSeries1.BarOffsetPercent := -30;
    chChartBarSeries2.BarWidthPercent := 25;
    chChartBarSeries2.BarOffsetPercent := 0;
    chChartBarSeries3.BarWidthPercent := 25;
    chChartBarSeries3.BarOffsetPercent := 30;
  end
  else if iNumDiag = 4 then
  begin
    chChartBarSeries1.BarWidthPercent := 20;
    chChartBarSeries1.BarOffsetPercent := -34;
    chChartBarSeries2.BarWidthPercent := 20;
    chChartBarSeries2.BarOffsetPercent := -11;
    chChartBarSeries3.BarWidthPercent := 20;
    chChartBarSeries3.BarOffsetPercent := 11;
    chChartBarSeries4.BarWidthPercent := 20;
    chChartBarSeries4.BarOffsetPercent := 34;
  end
  else if iNumDiag = 5 then
  begin
    chChartBarSeries1.BarWidthPercent := 15;
    chChartBarSeries1.BarOffsetPercent := -34;
    chChartBarSeries2.BarWidthPercent := 15;
    chChartBarSeries2.BarOffsetPercent := -17;
    chChartBarSeries3.BarWidthPercent := 15;
    chChartBarSeries3.BarOffsetPercent := 0;
    chChartBarSeries4.BarWidthPercent := 15;
    chChartBarSeries4.BarOffsetPercent := 17;
    chChartBarSeries5.BarWidthPercent := 15;
    chChartBarSeries5.BarOffsetPercent := 34;
  end;
  if cbComboDiag1.ItemIndex > -1 then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag1.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource1.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor1);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries1.Title := cbComboDiag1.Text;
    chChartBarSeries1.SeriesColor := clColor1;
    chChartBarSeries1.Active := True;
  end;
  if cbComboDiag2.ItemIndex > -1 then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag2.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource2.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor2);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries2.Title := cbComboDiag2.Text;
    chChartBarSeries2.SeriesColor := clColor2;
    chChartBarSeries2.Active := True;
  end;
  if cbComboDiag3.ItemIndex > -1 then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag3.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource3.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor3);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries3.Title := cbComboDiag3.Text;
    chChartBarSeries3.SeriesColor := clColor3;
    chChartBarSeries3.Active := True;
  end;
  if cbComboDiag4.ItemIndex > -1 then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag4.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource4.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor4);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries4.Title := cbComboDiag4.Text;
    chChartBarSeries4.SeriesColor := clColor4;
    chChartBarSeries4.Active := True;
  end;
  if cbComboDiag5.ItemIndex > -1 then
  begin
    for iRowGridStat := 1 to sgStatistic.RowCount - 1 do
    begin
      if sgStatistic.Cells[0, iRowGridStat] = cbComboDiag5.Text then
        Break;
    end;
    iBottomAx := 1;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        csChartSource5.Add(
          iBottomAx,
          StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]),
          sgStatistic.Cells[iColGridStat, 0],
          clColor5);
        Inc(iBottomAx);
      end;
    end;
    chChartBarSeries5.Title := cbComboDiag5.Text;
    chChartBarSeries5.SeriesColor := clColor5;
    chChartBarSeries5.Active := True;
  end;
  chChart.Legend.Visible := True;
  chChart.Visible := True;
end;

procedure TfmMain.CreateDiagramAllWordsBook;
var
  iColGridStat, iBottomAx, i: integer;
  clColor1: TColor;
  blSelBookmark: boolean;
begin
  // Create diagram of all words
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  if sgStatistic.RowCount = 0 then
  begin
    pcMain.ActivePage := tsStatistic;
    MessageDlg(msg062,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if clDiagBookmark.Count = 0 then
  begin
    MessageDlg(msg063,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  blSelBookmark := False;
  for i := 0 to clDiagBookmark.Count - 1 do
  begin
    if clDiagBookmark.Checked[i] = True then
    begin
      blSelBookmark := True;
    end;
  end;
  if blSelBookmark = False then
  begin
    MessageDlg(msg064,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  fmInput.Caption := dgc022;
  fmInput.lbTitle.Caption := dgc023;
  if fmInput.ShowModal = mrOK then
  begin
    stDiagTitle := fmInput.edTitle.Text;
  end;
  if stDiagTitle = '' then
  begin
    chChart.Title.Visible := False;
  end
  else
  begin
    chChart.Title.Text.Clear;
    chChart.Title.Text.Add(stDiagTitle);
    chChart.Title.Visible := True;
  end;
  clColor1 := $8080FF;
  csChartSource1.DataPoints.Clear;
  chChartBarSeries1.Title := '';
  iBottomAx := 1;
  for iColGridStat := 2 to sgStatistic.ColCount - 1 do
  begin
    if clDiagBookmark.Checked[iColGridStat - 2] = True then
    begin
      csChartSource1.Add(
        iBottomAx,
        StrToInt(sgStatistic.Cells[iColGridStat, sgStatistic.RowCount - 1]),
        sgStatistic.Cells[iColGridStat, 0],
        clColor1);
      Inc(iBottomAx);
    end;
  end;
  chChartBarSeries1.SeriesColor := clColor1;
  chChartBarSeries1.BarWidthPercent := 70;
  chChartBarSeries1.BarOffsetPercent := 0;
  chChartBarSeries1.Active := True;
  chChart.Legend.Visible := False;
  chChart.Visible := True;
end;

procedure TfmMain.CreateDiagramAllWordsNoBook;
var
  iRowGridStat, iColGridStat, iBottomAx, iTot, i: integer;
  clColor1: TColor;
  blSelBookmark: boolean;
begin
  // Create diagram of all words with no bookmarks
  chChart.Visible := False;
  chChartBarSeries1.Active := False;
  chChartBarSeries2.Active := False;
  chChartBarSeries3.Active := False;
  chChartBarSeries4.Active := False;
  chChartBarSeries5.Active := False;
  if sgStatistic.RowCount = 0 then
  begin
    pcMain.ActivePage := tsStatistic;
    MessageDlg(msg065,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  if clDiagBookmark.Count = 0 then
  begin
    MessageDlg(msg066,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  blSelBookmark := False;
  for i := 0 to clDiagBookmark.Count - 1 do
  begin
    if clDiagBookmark.Checked[i] = True then
    begin
      blSelBookmark := True;
    end;
  end;
  if blSelBookmark = False then
  begin
    MessageDlg(msg067,
      mtWarning, [mbOK], 0);
    Abort;
  end;
  fmInput.Caption := dgc022;
  fmInput.lbTitle.Caption := dgc023;
  if fmInput.ShowModal = mrOK then
  begin
    stDiagTitle := fmInput.edTitle.Text;
  end;
  if stDiagTitle = '' then
  begin
    chChart.Title.Visible := False;
  end
  else
  begin
    chChart.Title.Text.Clear;
    chChart.Title.Text.Add(stDiagTitle);
    chChart.Title.Visible := True;
  end;
  clColor1 := $8080FF;
  csChartSource1.DataPoints.Clear;
  chChartBarSeries1.Title := '';
  iBottomAx := 1;
  for iRowGridStat := 1 to sgStatistic.RowCount - 2 do
  begin
    iTot := 0;
    for iColGridStat := 2 to sgStatistic.ColCount - 1 do
    begin
      if clDiagBookmark.Checked[iColGridStat - 2] = True then
      begin
        iTot := iTot + StrToInt(sgStatistic.Cells[iColGridStat, iRowGridStat]);
      end;
    end;
    csChartSource1.Add(
      iBottomAx,
      iTot,
      sgStatistic.Cells[0, iRowGridStat],
      clColor1);
    Inc(iBottomAx);
  end;
  chChartBarSeries1.SeriesColor := clColor1;
  chChartBarSeries1.BarWidthPercent := 70;
  chChartBarSeries1.BarOffsetPercent := 0;
  chChartBarSeries1.Active := True;
  chChart.Legend.Visible := False;
  chChart.Visible := True;
end;

procedure TfmMain.DisableMenuItems;
begin
  // Disable menu items
  miFileNew.Enabled := False;
  miFileOpen.Enabled := False;
  miFileSave.Enabled := False;
  miFileSaveAs.Enabled := False;
  miFileSetBookmark.Enabled := False;
  miFileUpdBookmark.Enabled := False;
  miConcordanceCreate.Enabled := False;
  miFileOpenConc.Enabled := False;
  miFileSaveConc.Enabled := False;
  miConcodanceShowSelected.Enabled := False;
  miConcordanceAddSkip.Enabled := False;
  miConcordanceDelCont.Enabled := False;
  miConcordanceRemove.Enabled := False;
  miConcordanceJoin.Enabled := False;
  miConcordanceRefreshGrid.Enabled := False;
  miConcordanceOpenSkip.Enabled := False;
  miConcordanceSaveSkip.Enabled := False;
  miConcordanceSaveRep.Enabled := False;
  miStatisticCreate.Enabled := False;
  miStatisticSortName.Enabled := False;
  miStatisticSortFreq.Enabled := False;
  miStatisticSave.Enabled := False;
  miDiagramSingleWordsBook.Enabled := False;
  miDiagramAllWordNoBook.Enabled := False;
  miDiagramAllWordsBook.Enabled := False;
  miDiagramShowVal.Enabled := False;
  miDiagramShowGrid.Enabled := False;
  miDiaZoomIn.Enabled := False;
  miDiaZoomOut.Enabled := False;
  miDiaZoomNormal.Enabled := False;
  miDiagramSave.Enabled := False;
  meSkipList.ReadOnly := True;
  edFltStart.ReadOnly := True;
  edFltEnd.ReadOnly := True;
  edLocate.ReadOnly := True;
  bnDeselConc.Enabled := False;
end;

procedure TfmMain.EnableMenuItems;
begin
  // Enable menu items
  miFileNew.Enabled := True;
  miFileOpen.Enabled := True;
  miFileSave.Enabled := blTextModSave;
  miFileSaveAs.Enabled := True;
  miFileSetBookmark.Enabled := True;
  miFileUpdBookmark.Enabled := True;
  miConcordanceCreate.Enabled := True;
  miFileOpenConc.Enabled := True;
  miFileSaveConc.Enabled := True;
  miConcodanceShowSelected.Enabled := True;
  miConcordanceAddSkip.Enabled := True;
  miConcordanceDelCont.Enabled := True;
  miConcordanceRemove.Enabled := True;
  miConcordanceJoin.Enabled := True;
  miConcordanceRefreshGrid.Enabled := True;
  miConcordanceOpenSkip.Enabled := True;
  miConcordanceSaveSkip.Enabled := True;
  miConcordanceSaveRep.Enabled := True;
  miStatisticCreate.Enabled := True;
  miStatisticSortName.Enabled := True;
  miStatisticSortFreq.Enabled := True;
  miStatisticSave.Enabled := True;
  miDiagramAllWordNoBook.Enabled := True;
  miDiagramAllWordsBook.Enabled := True;
  miDiagramSingleWordsBook.Enabled := True;
  miDiagramShowVal.Enabled := True;
  miDiagramShowGrid.Enabled := True;
  miDiaZoomIn.Enabled := True;
  miDiaZoomOut.Enabled := True;
  miDiaZoomNormal.Enabled := True;
  miDiagramSave.Enabled := True;
  meSkipList.ReadOnly := False;
  edFltStart.ReadOnly := False;
  edFltEnd.ReadOnly := False;
  edLocate.ReadOnly := False;
  bnDeselConc.Enabled := True;
end;

end.
