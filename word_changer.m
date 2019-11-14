function [] = word_changer ()
button = questdlg('Welcome. In order to continue, click yes and choose the .docx file you want to overwrite');
b = strcmp (button,'Yes');
%%
if b ==1
    [FileName,PathName,FilterIndex] = uigetfile('../*.docx');
    path (PathName,path);
    Word = actxserver ('Word.application');
    Word.Visible = 0;
    set (Word, 'DisplayAlerts',0);
    Docs = Word.Documents;
    NameandPath = {PathName, FileName};
    NameandPath = strjoin (NameandPath,'');
    path (PathName,path);
    Doc = Docs.Open(NameandPath);
    selection = Word.Selection;
end
    
    %%
    %here we need to load a vocabulary
    addpath ('E:\','E:\ML WORD CHANGER');
T = readtable ('VOCAB.xlsx');
%%
[rows, columns] = size(T);
%%


for i=1:rows
    t1 = T{i,1};
    t2 = strjoin(t1);
    t3 = 'very';
    t4 = [t3,' ',t2];
    t5 = T{i,2};
    t6 = strjoin(t5);
    selection.Find.Execute(t4,0,0,0,0,0,1,1,0,t6,2,0,0,0,0);
    Doc.Save;
    
end

%%
    Doc.Save;
    Doc.Close;
    invoke(Word, 'Quit');
    delete (Word);
h = msgbox('Operation Completed');

