####TODO VANAF OEF 31
####TODO ADD EXCEL FILE TO THIS
####MAKE SCRIPT THAT FIXES CODE $S --> $s
####TODO SPLIT CODE IN MODULES
use warnings;
use Win32::OLE 'in';
use Win32::OLE::Variant;
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Math::BigInt;
my $computername  = ".";
my $namespace = "root/cimv2";
my $Locator=Win32::OLE->new("WbemScripting.SWbemLocator");
my $WbemServices = $Locator->ConnectServer($computername, $namespace);
$Win32::OLE::Warn = 3;



sub get_instances_from_class{
	my $Class = shift;
	return $Class->Instances_(wbemFlagUseAmendedQualifiers);
}

#todo replace this in lot of code
sub get_class{
	return $WbemServices->Get(shift,wbemFlagUseAmendedQualifiers);
}

#vb. Win32_LocalTime, Win32_DiskDrive en Win32_Product
#print_class_qualifiers_class("Win32_LocalTime");
sub print_class_qualifiers_class{
	my $Class = shift; 
	$Class = get_class($Class) if (ref($Class) ne "Win32::OLE");
	foreach $Qualifier (in $Class->Qualifiers_){
	    print_qualifier($Qualifier);
	}
}

sub print_class_qualifiers_instance{
	$Instance = shift;
	print "\n",$Instance->{Path_}->{relpath},"\n";
	foreach $Qualifier (in $Instance->Qualifiers_){ #geeft een kortere lijst van qualifiers
		print_qualifier($Qualifier);
	}
}

sub print_qualifier{
	my $Qualifier = shift;
	my $waarde=$Qualifier->Value;
		print "\t",$Qualifier->Name," (",ref($waarde) eq "ARRAY"  ? "Array=".join (",",@{$waarde}) : $waarde,")\n";
}


#does object have qualifier + is it true?
sub isSetTrue{
  my ($Object,$prop)=@_;
  return  $Object->{Qualifiers_}->Item($prop) && $Object->{Qualifiers_}->Item($prop)->{Value};
}

#show all subclasse
sub print_all_classes{
	print_sub_classes("",-2);
}

#show all subclasses under the given class
#print_classes_below_class("__EventGenerator");
sub print_classes_below_class{
	$class = shift;
	$class = get_class($class) if (ref($class) ne "Win32::OLE");
	print_sub_classes($class,-1);
}

#use print_all_classes/print_classes_below methods
sub print_sub_classes {
    my ($ClassName,$Level) = @_;
    $Level++;
    print "\n","\t" x $Level , $ClassName  if $ClassName;

    my $Instances = $WbemServices->SubClassesOf($ClassName, wbemQueryFlagShallow); #onmiddelijke subklassen
    print_sub_classes($_,$Level) foreach sort {uc($a) cmp uc($b)} map {$_->{Path_}->{RelPath}} in $Instances;
}

#print the size of the directory
#print_directory_size("c:\\oraclexe", 10);
sub print_directory_size{
	my ($DirectoryName,$ShowLevel) = @_;
	my $Directory = $WbemServices->Get("Win32_Directory='$DirectoryName'");
	directory_size_recursive($Directory,$ShowLevel, $ShowLevel);
}

#use print_directory_size method
sub directory_size_recursive {
   my ($Directory,$Level,$ShowLevel) = @_;
   my $Size = Math::BigInt->new();

   my $Query = "ASSOCIATORS OF {Win32_Directory='$Directory->{Name}'} WHERE AssocClass=CIM_DirectoryContainsFile";
   $Size += $_->{FileSize} foreach in $WbemServices->ExecQuery($Query);
   $Query = "ASSOCIATORS OF {Win32_Directory='$Directory->{Name}'} 
                WHERE AssocClass=Win32_SubDirectory Role=GroupComponent";
   $Size += directory_size_recursive($_, $Level-1, $ShowLevel) foreach in $WbemServices->ExecQuery($Query);

   printf "%12s%s%s\n", $Size,("\t" x ($ShowLevel-$Level+1)), $Directory->{Name} if $Level >= 0;
   return $Size;
}

#print_active_services();
sub print_active_services{
	my $ClassName = "Win32_Service";
	my $Instances = $WbemServices->ExecQuery("Select * From $ClassName Where State = 'Running'");
	@toonprop=qw (DisplayName Name Description Status StartMode StartName);

	foreach my $Instance (in $Instances) {
	    print "\n--------------------------------------------------------------\n";
	    foreach my $prop (@toonprop){
	        printf "\t%-15s : %s\n ",$prop, $Instance->{$prop} if $Instance->{$prop} ;
	    }
	}
}


#print_computer_power_logs();
sub print_computer_power_logs{
	#query om WMI-objecten op te halen
	my $datetime = Win32::OLE->new("WbemScripting.SWbemDateTime"); 
	$Query="Select * from Win32_NTLogEvent  Where Logfile = 'System'
                                       and ( EventCode = '6005' or EventCode = '6006' )
                                       and SourceName = 'EventLog'";
	$Instances = $WbemServices->ExecQuery($Query);
	foreach my $Instance (sort {$a->{TimeWritten} cmp $b->{TimeWritten}} in $Instances) {
	    $datetime->{Value} = $Instance->{TimeWritten} ;
	    $periode=($datetime->GetFileTime - $StartupTime)/10000000 if $StartupTime;
	    printf "%-22s: %s %s\n",$datetime->GetVarDate, $Instance->{Message} =~ /(started|stopped)/
	                ,($Instance->{EventCode} == 6006 && $StartupTime ? "na $periode s" : "");
	    $StartupTime=($Instance->{EventCode} == 6005 ? $datetime->GetFileTime : undef);
	}
}

#print the attributes+values from all instances of a class
#print_attributes_from_class_instances("Win32_Environment");
sub print_attributes_from_class_instances{
	my $class = shift;
	$class = get_class($class) if (ref($class) ne "Win32::OLE");
	my $instances = get_instances_from_class($class);
	print $instances->{Count} , " exempla(a)r(en) \n"; #er zal maar 1 exemplaar zijn
	foreach my $instance (in $instances){
		print "\n\nInstance:\n\n";
		foreach my $prop (in $instance->{Properties_}, $instance->{SystemProperties_}){
		    if ($prop->{CIMType} != 101){ #datum
		        printf "%-42s : %s %s \n",$prop->{Name},
		            ($prop->{Isarray} ? "(ARRAY)" : "",
		            ($prop->{Isarray} ? join ",",@{$prop->{Value}} : $prop->{Value}));
		    }
		}
	}

}

#print the attributes from the class with the given name 
#print_attributes_from_class("Win32_Environment");
sub print_attributes_from_class{
	my $class = shift;
	$class = get_class($class) if (ref($class) ne "Win32::OLE");
	
	%wd = %{Win32::OLE::Const->Load($Locator)}; 
	my %types;
	while (($type,$nr)=each (%wd)){
	  	if ($type=~/Cimtype/){
	    	$type=~s/wbemCimtype//g;
	    	$types{$nr}=$type;
	  	}
	}

	print  $class->{SystemProperties_}->{__CLASS}->{Value}," bevat ", $class->{Properties_}->{Count}," properties en ", 
					     $class->{SystemProperties_}->{Count}," systemproperties : \n\n";

	foreach my $prop (in $class->{Properties_}, $class->{SystemProperties_}){
	       print "\t",$prop->{Name}," (",$prop->{CIMType}, "/", $types{$prop->{CIMType}} , ($prop->{Isarray} ? " - is array" : ""),")\n";
	}	
}
#print the environment variables from the class with the given name ex "Win32_Environment"
#TODO fix this junk method
#print_environment_variables_classname("Win32_Environment");
sub print_environment_variables_classname{
	$classname = shift;
	printf  "%s=%s [%s] [%s]\n",$_->{Name},$_->{VariableValue},$_->{UserName},$_->{SystemVariable}
	foreach sort {uc($a->{Name}) cmp uc($b->{Name})} in $WbemServices->InstancesOf($classname); 		
}



#prints the objects linked to the given instance + option to show only the classes
# my $instance = $WbemServices->Get("Win32_Directory.Name='c:\\'");
# print_objects_associated_with($instance,1);
sub print_objects_associated_with{
	my $instance = shift;
	my $class_only = shift;
	my $associators;
	if(!$class_only){
		$associators = $instance->Associators_(); #alle instanties, geassocieerd met deze instantie	

	} else {
		$associators = $instance->Associators_(undef,undef,undef,undef,1); #enkel de klassen voor de geassocieerde objecten	
	}
	print $associators->{Count} , " exempla(a)r(en) \n";
}

#use get_namespaces method
sub recursive_get_namespaces {
	$name_space = shift;
	print $name_space , "\n";
	my $wbem_services = $Locator->ConnectServer($computername, $name_space);
	return if (Win32::OLE->LastError()); #indien geen connectie kan gemaakt worden met deze namespace
	my $Instances = $wbem_services->Execquery("select * from __NAMESPACE");
	return unless $Instances->{Count};
	#de naam van de Namespace moet worden opgebouwd, zodat het connecteren zal lukken.
	recursive_get_namespaces("$name_space/$_") foreach sort {uc($a) cmp uc($b)} map {$_->{Name}} in $Instances;
}

#get namespaces below given namespace
sub get_namespaces{
	$Win32::OLE::Warn = 0;
	recursive_get_namespaces(shift);
	$Win32::OLE::Warn = 3;
}

sub print_service_pack_info{
	$operating_inst = $WbemServices->Get("Win32_OperatingSystem=@"); #uniek instantie van de singleton-klasse

	#ServicePack-informatie
	print "ServicePackMajorVersion: " , $operating_inst->{ServicePackMajorVersion};
	print "\nServicePackMinorVersion: " , $operating_inst->{ServicePackMinorVersion};

	#Verdere informatie over de Windows-versie
	print "\nCaption: ",$operating_inst->{Caption};
	print "\nversion: ",$operating_inst->{Version},
	print "\nOSArchitecture: ",$operating_inst->{OSArchitecture};
}

#use start/stop_service method
sub change_service_state{
	my ($servicename, $action) = splice @_, 0,2 ;
	my $Instance = $WbemServices->Get("$ClassName='$servicename'");
	Win32::OLE->LastError()==0 || die Win32::OLE->LastError();
	printf "%s is currently %s\n" ,$Instance->{DisplayName},$Instance->{State};
	my $Methods = get_class($ClassName)->{Methods_};
	my %StartServiceReturnValues=makeHash_method_qualifier($Methods->Item("StartService"));
	my %StopServiceReturnValues=makeHash_method_qualifier($Methods->Item("StopService"));
	if ( $action eq "start" ) {
	   my $OutParameters = $Instance->ExecMethod_("StartService"); 
	   my $intRC = $OutParameters->{ReturnValue};
	   $intRC ? print "Execution failed: " . $StartServiceReturnValues{$intRC} ,"\n": print $Instance->{DisplayName} . " started\n";
	}
	else {
	   $intRC=$Instance->StopService();  
	   $intRC ? print "Execution failed: " . $StopServiceReturnValues{$intRC},"\n" : print $Instance->{DisplayName} . " stopped\n";
	}  
}

#start service with name
sub start_service{
	change_service_state(shift,"start");
}
#stop service with name
sub stop_service{
	change_service_state(shift,"stop");
}

#destroy every proces with name
sub destroy_processes_with_name{
	 $process_name = shift;
	 foreach (in $WbemServices->ExecQuery("Select * from Win32_Process Where Name='$process_name'")){
		$_->Terminate;
	 }
}

#create process with given exe path. returns the processhandles
#examples "notepad.exe", "C:\\Program Files (x86)\\Microsoft Office\\Office16\\excel.exe", "calc.exe","outlook.exe"
sub create_process{
	$process_name = shift;
	my $class = get_class("Win32_Process");
	my $methode = $class->{Methods_}->{"Create"};
	my %CreateReturnValues=makeHash_method_qualifier($methode);
	my $MethodInParameters =$methode->{InParameters}; 
	$MethodInParameters->{CommandLine}=$process_name;
	my $MethodOutParameters=$class->ExecMethod_("Create",$MethodInParameters) ;   
	$Return = $MethodOutParameters->{ReturnValue};
	print $process_name, " : ", $CreateReturnValues{$Return},"\n";
	if ($Return eq 0) {
		$Id = $MethodOutParameters->{ProcessId};
		my $relpad="Win32_Process.Handle=\"$Id\"";
		my $object =  $WbemServices->get($relpad);
		my $handle = $object->{Handle}; 
		return $handle;  
	}
	return -1;
}

#destroy process with given handle
sub destroy_process{
	$processHandle = shift;
	my $class = get_class("Win32_Process");
        my $method = $class->{Methods_}->{"Terminate"};
	my %TerminateReturnValues=makeHash_method_qualifier($method);
	my $MethodInParameters =$method->{InParameters}; #Moeten echter niet ingevuld worden
	my $relpad="Win32_Process.Handle=\"$processHandle\"";
	my $object =  $WbemServices->get($relpad);
	print "$relpad wordt gestopt : ";
	my $MethodOutParameters=$object->ExecMethod_("Terminate",$MethodInParameters);
	$Return = $MethodOutParameters->{ReturnValue};
	print $TerminateReturnValues{$Return},"\n";
}
  
#maps the ValueMap to the Values for a given method
sub makeHash_method_qualifier{
	my $Method=shift;
	my %hash=();
	@hash{@{$Method->Qualifiers_(ValueMap)->{Value}}} = @{$Method->Qualifiers_(Values)->{Value}};
	return %hash;
}

