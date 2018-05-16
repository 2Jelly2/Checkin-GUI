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
%初始化程序
handles.Judge={'Y';'N';'Now Checking'};
%定义Judge，存放三种出勤情况
handles.output = hObject;
guidata(hObject, handles);

function Open_Callback(hObject, ~, handles)
%定义菜单中的Open功能
handles.file=uigetfile('*.xls;*.xlsx');
%打开.xls,.xlsx后缀的文件
if ~isequal(handles.file,0)
    [handles.data,~,handles.raw] = xlsread(handles.file);
    %若文件正常读取，则将其data和raw写入相应数组
end
[handles.numStu,~]=size(handles.data);
%获得data行数（表内学生总数）
numT=strcat(num2str(handles.numStu),' Student(s) in Total');
set(handles.Total,'String',numT);
%将学生总数与其他字符串相连接，并在UI上显示
set(handles.Sheet,'Data',handles.raw(2:handles.numStu+1,2:5));
%将GUI表格所需内容从raw中写入
guidata(hObject, handles);

function Counter_CreateFcn(hObject, ~, ~)
%自动生成内容，在不同场合改变可编辑文字框的背景色
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function Counter_Callback(hObject, ~, handles)
%将文本框内容（要抽选的学生数量）存放到maxPick数组
handles.maxPick=str2double(get(hObject,'String'));
guidata(hObject,handles);

function Confirm_Callback(hObject, ~, handles)
handles.Store=randperm(handles.numStu,handles.maxPick);
%从学生总数中随机抽取规定的人数，并存放到Store中
handles.reading=single(1);
%定义计数器reading
Mark(handles,3);
%调用Mark函数（传递结构体handles和出勤检查状况st）
guidata(hObject,handles);

function Mark(handles,st)
%自建函数Mark，将各内容输出到UI
row=find(handles.data(:,1:1)==handles.Store(handles.reading));
%定义row，在data数组的序号列找到与Store中选中学生的序号列对应的元素的行数
Sheet=get(handles.Sheet,'Data');
%（为了方便读写）由结构体的Sheet中提取出仅供本函数调用的Sheet
Sheet(row:row,5:5)=handles.Judge(st);
set(handles.Sheet,'Data',Sheet);
%在Sheet中按照传递回的出勤状况在第五列“Attendance”中标记
if st~=3 && handles.reading+1<=handles.maxPick
%当Judge不为状态3（Checking），且计数器reading未超过最大值maxPick时
    handles.reading=handles.reading+1;
    %计数器reading自增1
    row=find(handles.data(:,1:1)==handles.Store(handles.reading));
    %获取下一个抽选到的学生在data中的行数
    Sheet(row:row,5:5)=handles.Judge(3);
    %使其Judge为状态3出勤状况（Checking）
    set(handles.Sheet,'Data',Sheet);
    %由本函数的Sheet输出到UI的结构体Sheet
end
rdNm=Sheet(row:row,3:3);
dhatd=strcat('Does <',rdNm,'> attend the class?');
set(handles.ifA,'String',dhatd);
%获取当前学生姓名，与字符串连接后在UI中显示
pRate=strcat('<',num2str(handles.reading),'/',num2str(handles.maxPick),'>');
set(handles.Process,'String',pRate);
%获取点名进度，与字符串连接后在UI中显示
name=Sheet(row:row,2:2);
imgName=strcat('','img/',num2str(name{1}),'.jpg','');
imshow(imgName);
%获取当前学生学号，计算出图片存储路径并显示

function varargout = exp2Ex_OutputFcn(~, ~, handles)
%系统自动生成函数
varargout{1} = handles.output;



function Yes_Callback(hObject, ~, handles)
%当学生出勤，调用Mark函数并传递出勤信息
Mark(handles,1);
if handles.reading+1<=handles.maxPick
%当计数器reading未达上限时，计数器自增1
    handles.reading=handles.reading+1;
end
guidata(hObject,handles);

function No_Callback(hObject, ~, handles)
%当学生缺勤，调用Mark函数并传递缺勤信息
Mark(handles,2);
if handles.reading+1<=handles.maxPick
%当计数器reading未达上限时，计数器自增1
    handles.reading=handles.reading+1;
end
guidata(hObject,handles);


function Save_Callback(~, ~, handles)
%定义菜单中的Save功能
OutputC=get(handles.Sheet,'Data');
%获取输出内容
OutputR=strcat('','B2:F',num2str(handles.numStu+1),'');
%获取输出范围
xlswrite(handles.file,OutputC,1,OutputR)
%写入Excel文档