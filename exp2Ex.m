function varargout = exp2Ex(varargin)
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @exp2Ex_OpeningFcn, ...
                   'gui_OutputFcn',  @exp2Ex_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end
if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end

function exp2Ex_OpeningFcn(hObject, ~, handles, varargin)
%��ʼ������
handles.Judge={'Y';'N';'Now Checking'};
%����Judge��������ֳ������
handles.output = hObject;
guidata(hObject, handles);

function Open_Callback(hObject, ~, handles)
%����˵��е�Open����
handles.file=uigetfile('*.xls;*.xlsx');
%��.xls,.xlsx��׺���ļ�
if ~isequal(handles.file,0)
    [handles.data,~,handles.raw] = xlsread(handles.file);
    %���ļ�������ȡ������data��rawд����Ӧ����
end
[handles.numStu,~]=size(handles.data);
%���data����������ѧ��������
numT=strcat(num2str(handles.numStu),' Student(s) in Total');
set(handles.Total,'String',numT);
%��ѧ�������������ַ��������ӣ�����UI����ʾ
set(handles.Sheet,'Data',handles.raw(2:handles.numStu+1,2:5));
%��GUI����������ݴ�raw��д��
guidata(hObject, handles);

function Counter_CreateFcn(hObject, ~, ~)
%�Զ��������ݣ��ڲ�ͬ���ϸı�ɱ༭���ֿ�ı���ɫ
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function Counter_Callback(hObject, ~, handles)
%���ı������ݣ�Ҫ��ѡ��ѧ����������ŵ�maxPick����
handles.maxPick=str2double(get(hObject,'String'));
guidata(hObject,handles);

function Confirm_Callback(hObject, ~, handles)
handles.Store=randperm(handles.numStu,handles.maxPick);
%��ѧ�������������ȡ�涨������������ŵ�Store��
handles.reading=single(1);
%���������reading
Mark(handles,3);
%����Mark���������ݽṹ��handles�ͳ��ڼ��״��st��
guidata(hObject,handles);

function Mark(handles,st)
%�Խ�����Mark���������������UI
row=find(handles.data(:,1:1)==handles.Store(handles.reading));
%����row����data�����������ҵ���Store��ѡ��ѧ��������ж�Ӧ��Ԫ�ص�����
Sheet=get(handles.Sheet,'Data');
%��Ϊ�˷����д���ɽṹ���Sheet����ȡ���������������õ�Sheet
Sheet(row:row,5:5)=handles.Judge(st);
set(handles.Sheet,'Data',Sheet);
%��Sheet�а��մ��ݻصĳ���״���ڵ����С�Attendance���б��
if st~=3 && handles.reading+1<=handles.maxPick
%��Judge��Ϊ״̬3��Checking�����Ҽ�����readingδ�������ֵmaxPickʱ
    handles.reading=handles.reading+1;
    %������reading����1
    row=find(handles.data(:,1:1)==handles.Store(handles.reading));
    %��ȡ��һ����ѡ����ѧ����data�е�����
    Sheet(row:row,5:5)=handles.Judge(3);
    %ʹ��JudgeΪ״̬3����״����Checking��
    set(handles.Sheet,'Data',Sheet);
    %�ɱ�������Sheet�����UI�Ľṹ��Sheet
end
rdNm=Sheet(row:row,3:3);
dhatd=strcat('Does <',rdNm,'> attend the class?');
set(handles.ifA,'String',dhatd);
%��ȡ��ǰѧ�����������ַ������Ӻ���UI����ʾ
pRate=strcat('<',num2str(handles.reading),'/',num2str(handles.maxPick),'>');
set(handles.Process,'String',pRate);
%��ȡ�������ȣ����ַ������Ӻ���UI����ʾ
name=Sheet(row:row,2:2);
imgName=strcat('','img/',num2str(name{1}),'.jpg','');
imshow(imgName);
%��ȡ��ǰѧ��ѧ�ţ������ͼƬ�洢·������ʾ

function varargout = exp2Ex_OutputFcn(~, ~, handles)
%ϵͳ�Զ����ɺ���
varargout{1} = handles.output;



function Yes_Callback(hObject, ~, handles)
%��ѧ�����ڣ�����Mark���������ݳ�����Ϣ
Mark(handles,1);
if handles.reading+1<=handles.maxPick
%��������readingδ������ʱ������������1
    handles.reading=handles.reading+1;
end
guidata(hObject,handles);

function No_Callback(hObject, ~, handles)
%��ѧ��ȱ�ڣ�����Mark����������ȱ����Ϣ
Mark(handles,2);
if handles.reading+1<=handles.maxPick
%��������readingδ������ʱ������������1
    handles.reading=handles.reading+1;
end
guidata(hObject,handles);


function Save_Callback(~, ~, handles)
%����˵��е�Save����
OutputC=get(handles.Sheet,'Data');
%��ȡ�������
OutputR=strcat('','B2:F',num2str(handles.numStu+1),'');
%��ȡ�����Χ
xlswrite(handles.file,OutputC,1,OutputR)
%д��Excel�ĵ�