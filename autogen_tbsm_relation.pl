#! c:/perl/bin/perl.exe -w

use strict;
use Win32::OLE;
use Win32::GUI();
use Cwd;
use File::Basename;

my $srcfile = ();

#Create the window and child controls.
my $mainwin = new Win32::GUI::Window (
    -pos         => [100, 100],
    -size        => [500, 300],
    -name        => "Window",
    -text        => "生成TBSM业务关系工具",
    -minsize		 => [500, 300],
    -maxsize     => [500, 300],
    -resizable   => 1,
   # -pushstyle   => WS_CLIPCHILDREN,
    #NEM Events for this window
   # -onResize    => \&MainResize,
    -onTerminate => sub {return -1;},    
);

$mainwin->AddButton (
    -name        => 'Find',
    -pos         => [30, 140],
    -size        => [100, 20],
    -text        => '选择文件',
    -onClick     => \&FindFiles,
    -align			 => 'center',
);


$mainwin->AddButton(
	-name	=> "Create",
	-pos  => [160,140],
	-size => [150, 20],
	-text => "生成影响关系文件",
	-align => 'center',
	-onClick => \&gentbsm,
);


$mainwin->AddTextfield(
	-name => 'tf',
	-pos => [20,30],
  -size => [450,100],
	-align => 'left',
	-number => 0,
	-readonly => 1,
);


$mainwin->AddTextfield(
	-name => 'display',
	-pos => [20,180],
  -size => [450,50],
	-align => 'left',
	-number => 0,
	-readonly => 1,
);


#show both windows and enter the Dialog phase.
$mainwin->Show();
Win32::GUI::Dialog();

# allow multiple files, only one filter
sub FindFiles{
my ( $file, @file);
my ( @parms );
push @parms,
  -filter =>
    ['xlsx - excel files', '*.xlsx',
     'xls - excel files', '*.xls',
     'All Files - *', '*'
    ],
  -directory => "c:\\program files",
  -title => 'Select a file';
	@file = Win32::GUI::GetOpenFileName ( @parms );

	#print "$_\n" for @file;
	$mainwin->tf->Text($file[0]);
	$srcfile = $file[0];
	
  return 0;
}

sub gentbsm{
	
my $src = $srcfile;
my $dest = "export.radsh";
my $path = getcwd();

#my @myfiles =  split(/\./,basename $src);
#$dest = $myfiles[0]."-".$dest;

my $res1 = "开始产生\"$dest\",请等待...\t";
my $res2 = "已经生成\"$dest\"\n";

print "$res1\n";


$mainwin->display->Text("");
$mainwin->display->Append(${res1});


#如果输出文件存在,先删除他
if(-e "${path}/${dest}"){
	unlink 	"${path}/${dest}";
}

#打开输出文件
open(DFD, ">", "$path/$dest") or die("can not open $path/$dest. err: $!");

#打开excel   
my $excel = Win32::OLE->GetActiveObject('Excel.Application')
    || Win32::OLE->new('Excel.Application', 'Quit'); 
$excel->{Visible} = 0;
$excel->{DisplayAlerts} = 'False';

#打开工作簿
my $book = $excel->Workbooks->Open("${src}") 
		|| die("Unable to open document ", Win32::OLE->LastError());
#打开表单
my $sheet = $book->Worksheets("梳理内容");
#得到使用行数
my $totalLine = $sheet->UsedRange->Rows->Count;


my($tmp, $j, $k, $para) = ();
my(@tmparr, @inamearr, @idisplayarr, @iiparr, @ivalarr, @ideparr ) = ();
my($tname, $tuni, $tf1, $tf2, $ttd, $iname, $idisplay, $iip, $ival, $idep) = ();
#创建tbsm模板
foreach my $row (3 .. $totalLine)
{
	#得到模板名
	$tmp = $sheet->Cells($row, 1)->{Value};
	@tmparr = split(/\n/, $tmp);
	$tmp = $tmparr[0];
	$tmp =~ /^\s*(.+)\s*$/;
	$tname = $1;

	#得到确定唯一实例字段组合
	$tmp = $sheet->Cells($row, 2)->{Value};
	@tmparr = split(/\n/, $tmp);
	for($j=0; $j<@tmparr; $j++){
		$tmp = $tmparr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$tmparr[$j] = $1;
	}
	$tuni = join(",", @tmparr);

	#得到影响字段
	$tmp = $sheet->Cells($row, 3)->{Value};
	@tmparr = split(/\n/, $tmp);
	for($j=0; $j<@tmparr; $j++){
		$tmp = $tmparr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$tmparr[$j] = $1;
	}
	if($tmparr[0] eq "NA"){
		$tf1 = "NA";
		$tf2 = "NA";
	}else{
		$tf1 = $tmparr[0];
		$tf2 = $tmparr[1];
	}	

	$ttd = "NA";
	$para = join("#", $tname, $tuni, $tf1, $tf2, $ttd);
	
	#下面这句用来调试,打印关键信息
	#print $para."\n";
	# template#AlertKey,NodeAlias#N_SubComponentId:Process,Port,Log,Connection#Severity#dep1,dep2
	&addFullTemplate($para);
}


#创建模板依赖关系
foreach my $row (3 .. $totalLine)
{

	#得到模板名
	$tmp = $sheet->Cells($row, 1)->{Value};
	@tmparr = split(/\n/, $tmp);
	$tmp = $tmparr[0];
	$tmp =~ /^\s*(.+)\s*$/;
	$tname = $1;

	#得到依赖模板
	$tmp = $sheet->Cells($row, 4)->{Value};
	@tmparr = split(/\n/, $tmp);
	for($j=0; $j<@tmparr; $j++){
		$tmp = $tmparr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$tmparr[$j] = $1;
	}
	#有依赖模板
	if($tmparr[0] ne "NA" ){
		for($j=0; $j<@tmparr; $j++){
			#传入参数:name#dep1 或者 dep2:p
			&addFullDepRule( join("#", $tname, $tmparr[$j]) );	
		}
	}
}


#创建所有tbsm实例
foreach my $row (3 .. $totalLine)
{
	#得到模板名
	$tmp = $sheet->Cells($row, 1)->{Value};
	@tmparr = split(/\n/, $tmp);
	$tmp = $tmparr[0];
	$tmp =~ /^\s*(.+)\s*$/;
	$tname = $1;

	#得到确定唯一实例字段组合
	$tmp = $sheet->Cells($row, 2)->{Value};
	@tmparr = split(/\n/, $tmp);
	for($j=0; $j<@tmparr; $j++){
		$tmp = $tmparr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$tmparr[$j] = $1;
	}
	$tuni = join(",", @tmparr);


	#得到上一步模板创建的实例数组
	$tmp = $sheet->Cells($row, 5)->{Value};
	@inamearr = split(/\n/, $tmp);
	for($j=0; $j<@inamearr; $j++){
		$tmp = $inamearr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$inamearr[$j] = $1;
	}

	#得到上一实例的中文显示名称
	$tmp = $sheet->Cells($row, 6)->{Value};
	@idisplayarr = split(/\n/, $tmp);
	for($j=0; $j<@idisplayarr; $j++){
		$tmp = $idisplayarr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$idisplayarr[$j] = $1;
	}	

	#得到上一实例的ip
	$tmp = $sheet->Cells($row, 7)->{Value};
	@iiparr = split(/\n/, $tmp);
	for($j=0; $j<@iiparr; $j++){
		$tmp = $iiparr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$iiparr[$j] = $1;
	}	

	#得到上一实例的实例值
	$tmp = $sheet->Cells($row, 8)->{Value};
	@ivalarr = split(/\n/, $tmp);
	for($j=0; $j<@ivalarr; $j++){
		$tmp = $ivalarr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$ivalarr[$j] = $1;
	}

	

		#处理一个实例一个ip	
		for($j=0; $j<@inamearr; $j++){
						
						$iname = $inamearr[$j];
						#相同的中文显示名
						if(@idisplayarr ==  1){
							$idisplay = $idisplayarr[0];
						}else{
							$idisplay = $idisplayarr[$j];
						}
				
						#相同的ip
						if(@iiparr ==  1){
							$iip = $iiparr[0];
						}else{
							$iip = $iiparr[$j];
						}
				
						#相同的实例值
						if(@ivalarr ==  1){
							$ival = $ivalarr[0];
						}else{
							$ival = $ivalarr[$j];
						}
				
						#第三行表示影响关系的第一层,也就是跟节点,根节点不显示ip
						$para = join("#", $tname, $iname, $idisplay, $iip, $tuni, $ival);
					
						#下面这句用来调试,打印关键信息
						#print $para."\n";
						# FZ_Market#FZ_Market1#行情#192.168.1.1#AlertKey,NodeAlias#wtserver,90,80,ms_offline
						&addFullInstance($para);
		}
}

my($father1, $child1, @childarr) = ();
#添加tbsm实例之间依赖关系
foreach my $row (3 .. $totalLine)
{
	#得到实例之间的依赖关系
	$tmp = $sheet->Cells($row, 9)->{Value};
	@ideparr = split(/\n/, $tmp);
	for($j=0; $j<@ideparr; $j++){
		$tmp = $ideparr[$j];
		$tmp =~ /^\s*(.+)\s*$/;
		$ideparr[$j] = $1;
	}

	#有实例依赖关系
	if($ideparr[0] ne "NA"){
		for($j=0; $j<@ideparr; $j++){
			$idep = $ideparr[$j];
			($father1, $child1) = split(/:/, $idep);
			@childarr = split(/,/, $child1);
			for($k=0; $k<@childarr; $k++){
				#调试使用
				#print $father1.":".$childarr[$k]."\n";
				#传入参数:father instance#child instance
				&addInstanceDep( join( "#", $father1, $childarr[$k]) );
			}
		}
	}
}

#关闭打开的工作簿
$book->Close;
#关闭输出文件
close(DFD);

print "$res2\n";
$mainwin->display->Append($res2);
  
return 0;

}

#############################################################################
###### 下面封装了rad_radshell支持的语法
###### 下面的程序都是基本程序,用来快速生成创建影响关系
#############################################################################

#创建一个模板
#传入参数: 模板名称
sub createTemp{
		my ($name) = @_;
		print DFD "createTemplate(
\"$name\", 
\"cloud_svg.gif\"
);";
	print DFD "\n\n";
}

#设置模板描述
#传入参数: 模板名称
sub setTempDes{
		my ($name) = @_;
		print DFD "setTemplateDescription(
\"$name\", 
\"$name Template\"
);";
	print DFD "\n\n";
}

#对于每一个模板设置模板管理用户
#这里默认设置为tipadmin
sub setTempAdmin{

#	
#	my($name) = @_;
#	print DFD "addUserPreferencesForTemplate(
#\"$name\", 
#\"PRIV_0:tbsmTemplateAdmin\", 
#\"tipadmin\"
#);";
#	print DFD "\n\n";

}

#在模板里创建incoming rule,并设置incoming rule名称
#选择数据源为omnibus的objectServer,处理的事件类型为ITM的事件和class(0)事件
#传入参数:模板名称#确定唯一实例字段数组(传入时根据具体情况决定传入几个)
#参数格式:name#e1,e2...
sub addTempSource{
	 my($all) = @_;
	 my($name, $arr) = split(/#/, $all);
	 my(@instanceUni) = split(/,/, $arr);
	 my $i;

	print DFD "addNewRawAttribute(
\"$name\", 
\"${name}"."StatusRule\", 
new String[] { \"IBM Tivoli Monitoring(87722)\",\"Default Class(0)\" },
new String[] { ";															

for($i=0; $i<@instanceUni; $i++){
	print DFD "\"$instanceUni[$i]\"";
  print DFD "," if( $i+1 < @instanceUni);
}

	print DFD "}, 
0, 
\"ObjectServer\"
);";
	print DFD "\n\n";	
}

#设置过滤器和过滤器阀值,这个只处理传入2个字段
#传入参数:模板名称,过滤器和影响规则
#传入参数eg: addTempFiltersAndThreshold("testA#N_SubComponentId,Connection,Severity,1,Bad");
sub addTempFiltersAndThreshold{
	my($all) = @_;
	my($name, $rule) = split(/#/, $all);
	my(@arrs) = split(/,/, $rule);

	print DFD "addRawAttributeThresholdSet(
\"$name\", 
\"$name"."StatusRule\", 
\"$arrs[4]\", 
null, 
new String[] { \"$arrs[0]\",\"$arrs[2]\" }, 
new String[] { \"=\",\">=\" }, 
new String[] { \"$arrs[1]\",\"$arrs[3]\" }, 
0
);";
	print DFD "\n\n";
}

#对于每一个incoming rule,设置2个过滤字段
#传入参数格式:name#filter:val1,val2..#filter
sub addTempFilters{
	my($all) = @_;
	my($name, $filter1, $filter2) = split(/#/, $all);
	my($f, $vals) = split(/:/, $filter1);
	my(@val) = split(/,/, $vals);
	my $i;

	for($i=0; $i<@val; $i++){
		#addTempFiltersAndThreshold("testA#N_SubComponentId,Connection,Severity,1,Bad");
		&addTempFiltersAndThreshold("${name}#${f},$val[$i],$filter2,1,Bad");
	}

	for($i=0; $i <@val; $i++){
		&addTempFiltersAndThreshold("${name}#${f},$val[$i],$filter2,0,Good");
	}
}

#对于每一个incoming rule设置 rule显示名
sub setStatusRuleDisName{
	my($name) = @_;
	print DFD  "addUserPreferencesForTemplate(
\"$name\", 
\"RuleDisplayName_$name"."StatusRule\", 
\"$name"."StatusRule\"
);";
	print DFD "\n\n";
}

#对于每一个incoming rule设置 rule描述
sub setStatusRuleDes{
	my($name) = @_;
	print DFD "addUserPreferencesForTemplate(
\"$name\", 
\"RuleDescription_$name"."StatusRule\", 
\"\"
);";
	print DFD "\n\n";
}

#对于每一个incoming rule设置KPI
sub setStatusRuleKPI{
	my($name) = @_;
	print DFD "addUserPreferencesForTemplate(
\"$name\", 
\"$name"."StatusRule_KPI\", 
\"true\"
);";
	print DFD "\n\n";
}

#创建一个完整的incoming rule
#传入参数eg: 模板名#确定唯一实例数组#影响字段和影响字段的值
#           name#e1,e2....#filter:val1,val2..#filter
sub addCompleteIncomingRule{
	my($all) = @_;
	my @elem = split(/#/, $all);
	my $name = $elem[0];

	&addTempSource( join("#", $name, $elem[1]));#定义唯一确定一个实例的字段
	&addTempFilters( join("#", $name, $elem[2], $elem[3]));#定义影响字段和影响程度
	&setStatusRuleDisName($name);
	&setStatusRuleDes($name);
	&setStatusRuleKPI($name);
}

#设置模板的依赖,每一个模板依赖需要设置显示名,描述,KPI
#设置模板依赖的显示名
#传入参数:模板,依赖
sub addDepDisName{
	my($all) = @_;
	my($name, $dep) = split(/#/, $all);
	print DFD "addUserPreferencesForTemplate(
\"$name\", 
\"RuleDisplayName_Dependency$dep\", 
\"\"
);";
	print DFD "\n\n";
}

#设置模板依赖的描述
#传入参数:模板,依赖
sub setDepDes{
	my($all) = @_;
	my($name, $dep) = split(/#/, $all);
	print DFD "addUserPreferencesForTemplate(
\"$name\", 
\"RuleDescription_Dependency$dep\", 
\"\"
);";
	print DFD "\n\n";
}

#设置模板依赖的KPI
#传入参数:模板,依赖
sub setDepKPI{
	my($all) = @_;
	my($name, $dep) = split(/#/, $all);
	print DFD "addUserPreferencesForTemplate(
\"$name\", 
\"Dependency${dep}_KPI\", 
\"true\"
);";
	print DFD "\n\n";
}

#设置单个依赖,这个函数用来设置多个子实例中,父模板依赖最坏那个子实例状态
#传入参数:模板,依赖
sub addWorstDep{
	my($all) = @_;
	my($name, $dep) = split(/#/, $all);
	print DFD "addWorstChildDependencyAttributeToTemplate(
\"$name\", 
\"$dep\", 
\"Dependency$dep\", 
\"Bad\", 
\"Marginal\", 
true
);";	
	print DFD "\n\n";
}

#设置单个依赖,这个函数用来设置多个子实例中,父模板依赖所有子实例状态的百分比
#默认70%,父模板红;1%,父模板黄,这是一个经验值
#传入参数:模板,依赖
sub addPercentDep{
	my($all) = @_;
	my($name, $dep) = split(/#/, $all);
	print DFD "addPercentageOfChildrenDependencyAttributeToTemplate(
\"$name\", 
\"$dep\", 
\"Dependency$dep\", 
\"All\", 
70, 
1, 
true, 
\"\", 
50
);";	
	print DFD "\n\n";
}

#创建一个完整的dep rule
#传入参数:name#dep1 或者 dep2:p 
#dep后面加P表示使用百分比依赖,默认使用最差依赖
sub addFullDepRule{
	my($all) = @_;
	my($name, $dep) = split(/#/, $all);
	my(@arr) = split(/:/, $dep);
	#使用最差依赖
	if(@arr == 1){
		&addDepDisName($all);
		&setDepDes($all);
		&setDepKPI($all);
		&addWorstDep($all);
	#使用百分比依赖	
	}else{
		&addDepDisName( join("#", $name, $arr[0]) );
		&setDepDes(join("#", $name, $arr[0]));
		&setDepKPI(join("#", $name, $arr[0]));
		&addPercentDep(join("#", $name, $arr[0]));
	}
}

#清除模板所有的ESDA
sub clearESDA{
	my($name) = @_;
print DFD "clearAllESDAsForTag(
\"$name\"
);";
	print DFD "\n\n";
}

#保存incoming rule 历史
sub saveStatuRuleHis{
	my($name) = @_;

	print DFD "addMetricMeta(
\"${name}.${name}StatusRule\", 
\"${name}StatusRule\", 
\"\", 
0, 
\"{\\\"type\\\":\\\"unspecified\\\"}\", 
0.0, 
100.0, 
900, 
true, 
false
);";
	print DFD "\n\n";
}

#保存依赖rule 历史
#参数:name#dep
sub saveDepHis{
	my($all) = @_;
	my($name, $dep) = split("#", $all);
	my(@arr) = split(":", $dep);
	$dep = $arr[0];

print DFD "addMetricMeta(
\"$name.Dependency${dep}Rule\", 
\"Dependency${dep}Rule\", 
\"\", 
0, 
\"{\\\"type\\\":\\\"unspecified\\\"}\", 
0.0, 
100.0, 
900, 
true, 
false
);";
	print DFD "\n\n";
}

#添加实例
#传入参数:name#instance#displau#ip
sub addInstance{
	my($all) = @_;
	my($name, $instance, $display, $ip) = split(/#/, $all);

	$display =	${display}.":".${ip}   if($ip ne "NA");
	print DFD "addServiceInstance(
new String[] { \"$name\" }, 
\"$instance\", 
\"${display}\", 
\"\", 
\"Standard\", 
null, 
null, 
\"REGULAR\"
);";

	print DFD "\n\n";
}


#清除实例所有已经配置的实例变量
#传入参数:instance
sub clearInstancePair{
	my($instance) = @_;
	print DFD "clearAllInstanceIDFieldValuePairs(
\"$instance\"
);";
	print DFD "\n\n";
}

#根据在模板中配置的唯一标识实例的字段,在这里添加字段对应的变量值
#传入参数:name#instance#field#val#index
sub addInstancePair{
	my($all) = @_;
	my($name, $instance, $field, $val, $index) = split(/#/, $all);
	print DFD "addInstanceIDFieldValuePair(
\"$name\", 
\"${name}StatusRule\", 
\"$field\", 
\"$val\", 
$index, 
\"$instance\"
);";
	print DFD "\n\n";
}

#实例配置GIS
#传入参数:instance
sub setInstanceGIS{
	my($instance) = @_;

	print DFD "addUserPreferencesForInstance(
\"$instance\", 
\"UseGIS\", 
\"false\"
);";
	print DFD "\n\n";
}

#配置实例经度
#传入参数:instance
sub setLatitude{
	my($instance) = @_;
	print DFD "addUserPreferencesForInstance(
\"$instance\", 
\"Latitude\", 
\"\"
);";
	print DFD "\n\n";
}

#配置实例显示顺序
#传入参数:instance
sub setOder{
	my($instance) = @_;
	print DFD  "addUserPreferencesForInstance(
\"$instance\", 
\"Order\", 
\"\"
);";
	print DFD "\n\n";
}

#配置实例显示玮度
#传入参数:instance
sub setLongitude{
	my($instance) = @_;
	print DFD "addUserPreferencesForInstance(
\"$instance\", 
\"Longitude\", 
\"\"
);";
	print DFD "\n\n";
}

#配置实例管理用户,默认设置为tipadmin
#传入参数:instance
sub  setInstanceUser{
#	my($instance) = @_;
#	print DFD "addUserPreferencesForInstance(
#\"$instance\", 
#\"PRIV_0:tbsmServiceAdmin\", 
#\"tipadmin\"
#);";
#	print DFD "\n\n";
}

#配置实例依赖关系
#传入参数:father instance#child instance
sub  addInstanceDep{
	my($all) = @_;
	my($father, $child) = split(/#/, $all);
	print	DFD "addServiceInstanceDependency(
\"$father\", 
\"$child\"
);";
	print DFD "\n\n";
}

#添加一个完整的模板
#传入参数格式:name#确定唯一实例字段组合#影响规则字段组合#依赖模板组合
#eg:
# template#AlertKey,NodeAlias#N_SubComponentId:Process,Port,Log,Connection#Severity#dep1,dep2
# 依赖模板可以没有,用NA表示
sub addFullTemplate{
	my($all) = @_;
	my($name, $uniStances, $filter1, $filter2, $deps) = split(/#/, $all);

	&createTemp($name);
	&setTempDes($name);

	#当模板的点是用来把所有实例点都汇合在他下面,他本身没有影响规则时,除了模板名变量有值,其他的都为NA
	if($uniStances ne "NA" and $filter1 ne "NA"  and $filter2 ne "NA"){
		#参数格式: name#e1,e2....#filter:val1,val2..#filter
		&addCompleteIncomingRule( join("#", $name, $uniStances, $filter1, $filter2 ) );
	}

	&setTempAdmin($name);
	&clearESDA($name);
	
	#&saveStatuRuleHis($name);
	#没有依赖模板
	#my $i;
	#if($dep[0] ne "NA" ){
	#	for($i=0; $i<@dep; $i++){
			#传入参数:name#dep1 或者 dep2:p
	#		saveDepHis( join("#", $name, $dep[$i]) );	
	#	}
	#}	
}

#添加一个完整的instance
#传入参数格式:name#instance#display#ip#field#value
#eg:
# FZ_Market#FZ_Market1#行情#192.168.1.1#AlertKey,NodeAlias#wtserver,90,80,ms_offline
sub addFullInstance{
	my($all) = @_;
	my($name, $instance, $display, $ip, $field, $value) = split(/#/, $all);
	my(@uniStances) = split(/,/, $field);
	my(@vals) = split(/,/, $value);


	&addInstance( join("#", $name, $instance, $display, $ip) );
	&clearInstancePair($instance);

	my $i;
	if($vals[0] ne "NA" and $ip ne "NA" ){
		for($i=0; $i<@vals; $i++){
				&addInstancePair( join("#", $name, $instance, $uniStances[0], $vals[$i],  $i+1) );
				&addInstancePair( join("#", $name, $instance, $uniStances[1], $ip, $i+1) );
		}
	}
	
	&setInstanceGIS($instance);
	&setLatitude($instance);
	&setLongitude($instance);
	&setOder($instance);
	&setInstanceUser($instance);
}
