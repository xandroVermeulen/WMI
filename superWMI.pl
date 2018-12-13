####TODO cim_managedsystemelemnt bevat coole klasses. vooral in logicalelement>logicaldevice
use warnings;
use Win32::OLE qw(EVENTS in);
use Win32::Console;
use Win32::OLE::Variant;
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Math::BigInt;
my $computername  = ".";
my $namespace = "root/cimv2";
my $Locator=Win32::OLE->new("WbemScripting.SWbemLocator");
my $WbemServices = $Locator->ConnectServer($computername, $namespace);
$Win32::OLE::Warn = 3;
use File::Spec;
open my $save_err, ">&STDERR";

my $typeLib=Win32::OLE::Const->Load($WbemServices);
my %cimtype;
while ( ( $key, $value ) = each %{$typeLib} ) {
	$cimtype{$value}=substr($key,11) if ($key=~/wbemCim/);
}
%wd = %{Win32::OLE::Const->Load($Locator)}; 

my %types;
while (($type,$nr)=each (%wd)){
	if ($type=~/Cimtype/){
	    $type=~s/wbemCimtype//g;
	    $types{$nr}=$type;
	}
}

if(@ARGV){
	handle_input();
}
###### examenvraag

sub thread_monitoring{
	excel_init();
	my $book = excel_open_file("threads.xlsx");
	my %threadCount=();
	foreach $proc (in $WbemServices->ExecQuery("select * from Win32_Process")){
		my $threads = $WbemServices->ExecQuery("select * from Win32_Thread where ProcessHandle = ".$proc->{Handle});
		$threadCount{$proc->{Name}} += $threads->{count};
	}
	my $sheet = $book->Sheets(1);
	my $mat;
	my $counter=1;
	while (($key,$value) = each(%threadCount)){
		$mat->[$counter][0] = $key;
		$mat->[$counter][1] = $value;
		$counter++;
	}
	my $range = $sheet->Range("A1:B".$counter);
	$range->{Value} = $mat;
	print "Press ENTER to exit"; <STDIN>;
}




##permanent eventregistration (only works on windows server)#################################################################

#time in ms
sub create_interval_timer{
	my ($name, $time) = @_;
	my $Instance = $WbemServices->Get("__IntervalTimerInstruction")->SpawnInstance_();
	$Instance->{TimerID}  = $name;
	$Instance->{IntervalBetweenEvents} = $time;
	$instancePath=$Instance->Put_(wbemFlagUseAmendedQualifiers);
	return $instancePath;
}
#timestring ex: 09/11/2018 11:00:00
sub create_oneshot_timer{
	#of op bepaald moment 1x triggeren
	my ($name, $time_string) = @_;
	my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");
	$Instance = $WbemServices->Get("__AbsoluteTimerInstruction")->SpawnInstance_();
	$Instance->{TimerID}  = $name;
	$DateTime->SetVardate($time_string);
	$Instance->{EventDateTime} = $DateTime->{Value};
	$instancePath=$Instance->Put_(wbemFlagUseAmendedQualifiers);
	return $instancePath;
}
 
sub create_event_filter{
	my ($name, $timer_id, $query) = @_;
	my $InstanceEvent = $WbemServices->Get("__EventFilter")->SpawnInstance_();
	$InstanceEvent->{Name}=$name;
	$InstanceEvent->{QueryLanguage} = "WQL";
	if($query){
		$InstanceEvent->{Query} = $query;
	} else {
		$InstanceEvent->{Query} = "SELECT * FROM __TimerEvent where TimerID = '".$timer_id."'";
	}
	$Filter = $InstanceEvent->Put_(wbemFlagUseAmendedQualifiers);
	return $Filter->{path};
}



#message = "timer wordt gelogd op tijdstip %TIME_CREATED% op toestel %__SERVER%";
sub create_commandline_event_consumer{
	my ($name, $message) = @_;
	my $InstanceReaction = $WbemServices->Get("CommandLineEventConsumer")->SpawnInstance_();
	$InstanceReaction->{Name}=$name;
	$InstanceReaction->{CommandLineTemplate} =  "msg console /Time:5 ".$message;
	$Consumer = $InstanceReaction->Put_(wbemFlagUseAmendedQualifiers);
	return $Consumer->{path};  
}
# filename = 'C:\\\\temp\\log.txt'
sub create_logfile_event_consumer{
	my ($name, $filename, $message) = @_;
	my $InstanceReaction= $WbemServices->Get("LogFileEventConsumer")->SpawnInstance_();
	$InstanceReaction->{Name}=$name;
	$InstanceReaction->{FileName} = $filename;
	$InstanceReaction->{Text} = "timer wordt gelogd op tijdstip $message";
	$Consumer = $InstanceReaction->Put_(wbemFlagUseAmendedQualifiers);
	return $Consumer->{path};  
}

sub create_smptpmail_event_consumer{
	my ($name, $from, $to, $subject, $server) = @_;
	my $InstanceReaction= $WbemServices->Get("SMTPEventConsumer")->SpawnInstance_();
	$InstanceReaction->{Name}=$name;
	$InstanceReaction->{FromLine}=$from;
	$InstanceReaction->{ToLine}=$to;
	$InstanceReaction->{Subject}=$subject;
	$InstanceReaction->{SMTPServer}=$server;
	$Consumer = $InstanceReaction->Put_(wbemFlagUseAmendedQualifiers);
	return $Consumer->{path};  
}



sub create_filter_to_consumer_binding{
	my ($filterpath, $consumerpath) = @_;
	my $InstanceCoupling = $WbemServices->Get("__FilterToConsumerBinding")->SpawnInstance_();
	$InstanceCoupling->{Filter}   = $filterpath; 
	$InstanceCoupling->{Consumer} = $consumerpath;
	$Result=$InstanceCoupling->Put_(wbemFlagUseAmendedQualifiers);
	print "\n1.\n",$Result->{Path},"\n";
	return $Result;
}
#create_event_registration("name", "interval", 6000, "logfile", "timer wordt gelogd op tijdstip %TIME_CREATED% op toestel %__SERVER%", 'C:\\\\temp\\log.txt');
sub create_event_registration{
	my ($name, $timertype, $time, $consumertype, $message, $filename) = @_;
	my ($timerpath, $consumerpath, $filterpath, $result);

	$timerpath = create_interval_timer($name, $time) if($timertype eq "interval");
	$timerpath = create_oneshot_timer($name, $time) if($timertype eq "oneshot");

	$filterpath = create_event_filter($name."_filter", $name);

	$consumerpath = create_commandline_event_consumer($name, $message) if($consumertype eq "commandline");
	$consumerpath = create_logfile_event_consumer($name, $message, $filename) if($consumertype eq "logfile");
	
	$result = create_filter_to_consumer_binding($filterpath, $consumerpath);
}

sub delete_all_event_registration_objects{
	my $allInstances = $WbemServices->InstancesOf("__IndicationRelated");
	print $allInstances->{Count}," instanties worden nu verwijderd\n";
	$_->Delete_() foreach in $allInstances;
	print Win32::OLE->LastError();
}

sub temp_method_ex59{
	$filterp = create_event_filter("test2", undef, "SELECT * FROM __InstanceCreationEvent 
	                     WITHIN 10 
	                     WHERE TargetInstance ISA 'Win32_Process' 
	                     AND (TargetInstance.Name = 'notepad.exe'
	                     OR  TargetInstance.Name = 'calc.exe')");

	$cpath = create_smptpmail_event_consumer(
		"test2", "marlemen.denert\@ugent.be", "marlemen.denert\@ugent.be",
		"%TargetInstance.Caption% started on %TargetInstance.__SERVER%", "smtp.ugent.be");
	create_filter_to_consumer_binding($filterp, $cpath);
}

#sync monitoring#######################################################################################################################
#monitor all services on the pc
sub watch_services{
	$Win32::OLE::Warn = 0;
	my $notif_query = "SELECT * FROM __InstanceModificationEvent WITHIN 5 WHERE TargetInstance ISA 'Win32_Service'";
	my $event_notif = $WbemServices->ExecNotificationQuery($notif_query);
	$|=1;
	print "Waiting for events: ";
	while(1){
		my $event = $event_notif->NextEvent(5000);
		Win32::OLE->LastError() and print "." or printf "\n%s changed from %s to %s\n"
			,$event->{TargetInstance}->{DisplayName}
			,$event->{PreviousInstance}->{State}
			,$event->{TargetInstance}->{State};
	}
	$Win32::OLE::Warn = 3;
}

sub watch_service{
	print "ToDo";
}

sub watch_processes{
	my $query  =  "SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process' ";
	my $EventSource = $WbemServices->ExecNotificationQuery($query);
	print "Watching processes:\n";
	while (1) {
	    my $Event = $EventSource->NextEvent(4000);
	    my $className=$Event->{Path_}->{Class};
	    next if  $className eq "__InstanceModificationEvent";   #enkel aanpassingen worden niet opgevolgd
	    printf "Process %-29s (%s)started \n", $Event->{TargetInstance}->{Name}, 
	                  $Event->{TargetInstance}->{Handle} if  $className eq "__InstanceCreationEvent";
	    printf "Process %-29s (%s)stopped \n", $Event->{TargetInstance}->{Name},
	                  $Event->{TargetInstance}->{Handle} if  $className eq "__InstanceDeletionEvent";
	}
}

sub watch_process{
	print "ToDo";
}
#async monitoring######################################################################
#console werkt niet
sub watch_processes_async{
	my $Sink = Win32::OLE->new ('WbemScripting.SWbemSink');
	Win32::OLE->WithEvents($Sink,\&EventCallBack);  #koppel de gewenste methode aan dit object

	my $Query1 = "SELECT * FROM __InstanceOperationEvent WITHIN 5 WHERE TargetInstance ISA 'Win32_Process'";
	$WbemServices->ExecNotificationQueryAsync($Sink, $Query1); 

	my $Console  = new Win32::Console(STD_INPUT_HANDLE);  #creeert een nieuw Console object
	$Console->Mode( ENABLE_PROCESSED_INPUT);    #enkel reageren op toetsen, niet op muis-bewegingen

	until ($Console->Input()) {   #zolang er geen input is
	 	Win32::OLE->SpinMessageLoop();
		Win32::Sleep(500);
	}

	$Sink->Cancel();
	Win32::OLE->WithEvents($Sink); #geen afhandeling meer bij dit SinkObject
}

#methode die de events afhandelt
sub EventCallBack(){
	my ($Source,$EventName,$Event,$Context) = @_;
	return unless $EventName eq "OnObjectReady";
	my $className=$Event->{Path_}->{Class};
	return if  $className eq "__InstanceModificationEvent";
	    
	if ($Event->{TargetInstance}->{Path_}->{Class} eq "Win32_Process") {
	    printf "%-29s started\n", $Event->{TargetInstance}->{Name} if  $className eq "__InstanceCreationEvent";
	    printf "%-29s stopped\n", $Event->{TargetInstance}->{Name} if  $className eq "__InstanceDeletionEvent";
	} else {
	    printf "%-29s %s\n", $Event->{TargetInstance}->{Path_}->{Class}, $Event->{TargetInstance}->{Path};
	}
}




#general###############################################################################
sub show_usb_devices{
	my $Query = 'select * from Win32_USBControllerDevice';
	my $Instances = $WbemServices->ExecQuery($Query);
	print_attributes_from_instances($Instances);
}

sub say_out_loud{
	my ($message) = @_;
	my $tts = Win32::OLE->new("SAPI.SpVoice");
	$tts->Speak($message);
} 

sub show_files_current_directory{
	show_files_directory(".");
}

sub show_files_directory{
	my ($name) = @_;
	my $fso = Win32::OLE->new("Scripting.FileSystemObject");
	my $folder = $fso->getFolder($name);
	foreach (in $folder->{Files}){
		printf "%-20s: %s\n", $_->{Name}, $_->{Type};
	}
}



sub handle_input{
	$method = shift @ARGV;
	&$method(@ARGV);
}

#only works in admin mode
sub create_environment_variable{
	my ($name, $value) = @_;
	my $class = get_class("Win32_Environment");
	my $instance = $class->SpawnInstance_(); #static method
	$instance->{Username} = "<SYSTEM>";
	$instance->{SystemVariable} = 1;
	$instance->{Name} = $name;
	$instance->{VariableValue} = $value;
	my $new_instance_path = $instance->Put_();
	print "Succes!\n" unless Win32::OLE->LastError();
	print "Return = ", $new_instance_path->{Path}, "\n";
	print "Return = ", $new_instance_path->{RelPath},"\n";
}

sub delete_environment_variable{
	my $name = shift; 
	$WbemServices->Get("Win32_Environment.UserName='<SYSTEM>',Name='$name'")->Delete_;
}

# sub temp{
# 	my $class = get_class("Win32_Process");
# 	my @methods = in $class->{Methods_};
# 	foreach $met (@methods){
# 		if($met->{InParameters}){
# 			foreach $prop (in $met->{OutParameters}->{properties_}){
# 				print $prop->{Name},"\n";
# 			}
# 		}

# 	}
# }

# sub temp2{
# 	my $class = get_class("Win32_Process");
# 	my @instances = in get_instances_from_class($class);
# 	my $i = $instances[0];
# 	print $i->{Properties_}->{Caption}->{Value};
# 	print "\n\n";
# 	print $i->{Caption};
# }

# sub temp3{
# 	my $class = get_class("Win32_Process");
# 	my $instances = get_instances_from_class($class);
# 	foreach $in (in $instances){
# 		my $props = $in->{Properties_};
# 		foreach $p (in $props){
# 			print $p->{Name};
# 		}
# 		print "\n";
# 	}
# }

sub print_method_input_parameters{
	my $method = shift;
	my @inParams = in $method->{InParameters}->{properties_};
	foreach (@inParams){
		print "[" if isSetTrue($_,"Optional");
		printf " %s(%s) ", $_->{Name}, $types{$_->{CIMTYPE}};
		print "]" if isSetTrue($_,"Optional");
	}
}

#print_methods("Win32_Share");
#create_object("Win32_Share");
sub create_object{
	my $class = get_class(shift);
	my $method = get_createby_method($class);
	my $inParams = $method->{InParameters}->{properties_};
	foreach (in $inParams){
		print "Optional: " if isSetTrue($_,"Optional");
		printf " %s(%s): ", $_->{Name}, $types{$_->{CIMTYPE}};
		$input = <STDIN>;
		chomp $input;
		$inParams->{$_->{Name}} = $input if ($input ne "");
	}
	$intRC = $class->ExecMethod_("Create", $inParams);
}

sub get_createby_method{
	$Class = shift;
	$Class = get_class($Class) if (ref($Class) ne "Win32::OLE");
    my $methodCreate = $Class->{Qualifiers_}->Item("CreateBy")->{Value};
    return $Class->{Methods_}->Item($methodCreate);
}

# print_register("");
# print_register("SYSTEM");
# print_register("SYSTEM\\Currentcontrolset\\Services\\tcpip");
sub print_register{
	my $start = shift;
	my %RootKey = ( HKEY_CLASSES_ROOT   => 0x80000000
	              , HKEY_CURRENT_USER   => 0x80000001
	              , HKEY_LOCAL_MACHINE  => 0x80000002
	              , HKEY_USERS          => 0x80000003
	              , HKEY_CURRENT_CONFIG => 0x80000005
	              , HKEY_DYN_DATA       => 0x80000006 );

	my $Registry = get_class("StdRegProv");

	my $InParameters=$Registry->{Methods_}->{EnumKey}->{InParameters};
	$InParameters->{Properties_}->Item(hDefKey)->{Value} = $RootKey{HKEY_LOCAL_MACHINE};

	print "$start";
	show_registry_branch($start,"",0,$Registry, $InParameters);
}

#use print_register method
sub show_registry_branch{
    my ($Key,$PrintNaam,$Level,$Registry,$InParameters)=@_;
   
    printf "%s%s\n",("\t" x $Level),$PrintNaam;

    $InParameters->{Properties_}->Item(sSubKeyName)->{Value} = $Key;
    my $EnumKeyOutParameters = $Registry->ExecMethod_(EnumKey,$InParameters);
    return if $EnumKeyOutParameters->{ReturnValue};
    if($EnumKeyOutParameters->{sNames}){
	    foreach my $SubKey (sort {lc($b) cmp lc($b)} @{$EnumKeyOutParameters->{sNames}}) {
	        my $Key = $Key.($Key ne "" ?"\\":"").$SubKey;
	        show_registry_branch($Key,$SubKey,$Level+1,$Registry,$InParameters) if $SubKey ne "";
	    }
	}
}

#print_methods("Win32_Volume");
sub print_methods{
	my $Class =  shift ;
	$Class = get_class($Class) if (ref($Class) ne "Win32::OLE");
	my $Methods = $Class->{Methods_};
	printf "\n%s bevat %d methodes met volgende aanroep (en extra return-waarde):\n ", 
	             $Class->{Path_}->{Class},$Methods->{Count};
	foreach my $Method (sort {uc($a->{Name}) cmp uc($b->{Name})} in $Methods) {
	    printf "\n\n\n------Methode %s ---(%s)-------------------- " ,$Method->{Name}, 
	             $Method->{Qualifiers_}->{Count};
	    foreach $qual (in $Method->{Qualifiers_}){    
	        printf "\n%s: %s\n",$qual->{name},
	                ref $qual->{value} eq "ARRAY" ? join " , ",@{$qual->{Value}} : $qual->{Value};
	    }    
	}
}

sub turn_off_errors{
	open STDERR, '>', File::Spec->devnull() or die "could not open STDERR: $!\n";
}

sub turn_on_errors{
	open STDERR, ">&", $save_err;
}

#print_disk_partition_info();
sub print_disk_partition_info{
	my $Class = get_class("Win32_LogicalDisk");

	my @DriveType = @{$Class->Properties_(DriveType)->Qualifiers_(Values)->{Value}}; #onthouden 
	my @MediaType_LogicalDrive = @{$Class->Properties_(MediaType)->Qualifiers_(Values)->{Value}}; #onthouden

	$Class = get_class("Win32_DiskDrive");
	my @Capabilities = @{$Class->Properties_(Capabilities)->Qualifiers_(Values)->{Value}};

	# MediaType heeft zowel Values als ValueMap - mechanisme 
	%MediaType_DiskDrive = make_property_hash("MediaType","Win32_DiskDrive");

	my $Instances = get_instances_from_class("Win32_DiskPartition");

	foreach $Instance (in $Instances) {
	    printf "*************************************** %s: %s\n", "DeviceID", $Instance->{DeviceID};
	    my $Properties = $Instance->{Properties_};

	    defined $_->{Value}
	       && printf "%s: %s\n",$_->{Name}, ref($_->{Value}) eq "ARRAY" ? join(",",@{$_->{Value}}) : $_->{Value}
	            foreach in $Properties;

	    print "\n";
	    $Query="ASSOCIATORS OF {Win32_DiskPartition='$Instance->{DeviceID}'} 
	            WHERE AssocClass=Win32_DiskDriveToDiskPartition";

	    foreach $PhysicalDiskInstance (in $WbemServices->ExecQuery($Query)) {
	        my $Properties = $PhysicalDiskInstance->{Properties_};  #haalt ook de waarden op
	        foreach $Property (in $Properties) {
	            if  ($Property->{Name} eq "Capabilities") {
	            	printf "%s: %s\n",$Property->{Name}, join(",",map {$Capabilities[$_]} @{$Property->{Value}});
	            }
	            elsif  ($Property->{Name} eq "MediaType") {
	                $value=$Property->{Value};
	                $value=~s/[\t]/ /g;  #tab-teken in plaats van een spatie !!
	             	printf "%s: (%s) %s\n", $Property->{Name},$Property->{Value},$MediaType_DiskDrive{$value};
	           }
	            else {
	                printf "%s: %s\n", $Property->{Name}, 
	                      ref( $Property->{Value} ) eq "ARRAY" ? join(",",@{$Property->{Value}}) : $Property->{Value}
	                   if defined $Property->{Value};
	            }
	        }
	    }

	    print "\n";
	    $Query="ASSOCIATORS OF {Win32_DiskPartition='$Instance->{DeviceID}'} 
	               WHERE AssocClass=Win32_LogicalDiskToPartition";
	    foreach $LogicalDiskInstance (in $WbemServices->ExecQuery($Query)) {
	        my $Properties = $LogicalDiskInstance->{Properties_};   #haalt ook de waarden op
	        foreach $Property (in $Properties) {
	            if  ($Property->{Name} eq "DriveType") {
	            	printf "%s: %s\n", $Property->{Name}, $DriveType[$Property->{Value}];

	            }
	            elsif  ($Property->{Name} eq "MediaType") {
	            	printf "%s: %s\n", $Property->{Name}, $MediaType_LogicalDrive[$Property->{Value}];
	            }
	            else {
	                printf "%s: %s\n", $Property->{Name}, 
	                      ref( $Property->{Value} ) eq "ARRAY" ? join ",",@{$Property->{Value}} : $Property->{Value}
	                   if defined $Property->{Value};
	            }
	        } 
	    }
	}
}

#print_network_adapter_info();
sub print_network_adapter_info{
	turn_off_errors();
	#op klassen de hash initialiseren
	%Availability        = make_property_hash("Availability","Win32_NetworkAdapter");
	%NetConnectionStatus = make_property_hash("NetConnectionStatus","Win32_NetworkAdapter");

	my $Query="SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionStatus>=0"; 
	my $AdapterInstances = $WbemServices->Execquery($Query);   #(*)

	foreach $AdapterInstance (sort {uc($a->{NetConnectionID}) cmp uc($b->{NetConnectionID})} in $AdapterInstances) {
	    print "******************************************************** \n";
	    printf "%s: %s\n", "Connection Name", $AdapterInstance->{NetConnectionID};

	    printf "%s: %s\n", "Adapter name", $AdapterInstance->{Name};
	    printf "%s: %s\n", "Device availability", $Availability{$AdapterInstance->{Availability}};
	    printf "%s: %s\n", "Adapter type", $AdapterInstance->{AdapterType};
	    printf "%s: %s\n", "Adapter state", $NetConnectionStatus{$AdapterInstance->{NetConnectionStatus}};
	    printf "%s: %s\n", "MAC address", $AdapterInstance->{MACAddress};
	    printf "%s: %s\n", "Adapter service name", $AdapterInstance->{ServiceName};
	    printf "%s: %s\n", "Last reset", $AdapterInstance->{TimeOfLastReset};

	    #Recource Informatie
	    $Query="ASSOCIATORS OF {Win32_NetworkAdapter='$AdapterInstance->{Index}'} 
	               WHERE AssocClass=Win32_AllocatedResource";
	    my $AdapterResourceInstances = $WbemServices->ExecQuery ($Query);
	    foreach $AdaptResInstance (in $AdapterResourceInstances) {
	        my $className=$AdaptResInstance->{Path_}->{Class};
	        printf "%s: %s\n", "IRQ resource", $AdaptResInstance->{IRQNumber} if $className eq "Win32_IRQResource";
	        printf "%s: %s\n", "DMA channel", $AdaptResInstance->{DMAChannel} if $className eq "Win32_DMAChannel";
	        printf "%s: %s\n", "I/O Port", $AdaptResInstance->{Caption}       if $className eq "Win32_PortResource";
	        printf "%s: %s\n", "Memory address", $AdaptResInstance->{Caption} if $className eq "Win32_DeviceMemoryAddress";
	    }

	    my $AdapterInstance = $WbemServices->Get ("Win32_NetworkAdapterConfiguration=$AdapterInstance->{Index}");
	    next unless $AdapterInstance->{IPEnabled};

	    if ($AdapterInstance->{DHCPEnabled}) {
	       printf "%s: %s\n", "DHCP expires", $AdapterInstance->{DHCPLeaseExpires};
	       printf "%s: %s\n", "DHCP obtained", $AdapterInstance->{DHCPLeaseObtained};
	       printf "%s: %s\n", "DHCP server", $AdapterInstance->{DHCPServer};
	    }

	    printf "%s: %s\n", "IP address(es)", (join ",",@{$AdapterInstance->{IPAddress}});
	    printf "%s: %s\n", "IP mask(s)", (join ",",@{$AdapterInstance->{IPSubnet}});
	    printf "%s: %s\n", "IP connection metric", $AdapterInstance->{IPConnectionMetric};
	    printf "%s: %s\n", "Default Gateway(s)",(join ",",@{$AdapterInstance->{DefaultIPGateway}});
	    printf "%s: %s\n", "Dead gateway detection enabled", $AdapterInstance->{DeadGWDetectEnabled};

	    printf "%s: %s\n", "DNS registration enabled", $AdapterInstance->{DomainDNSRegistrationEnabled};
	    printf "%s: %s\n", "DNS FULL registration enabled", $AdapterInstance->{FullDNSRegistrationEnabled};
	    printf "%s: %s\n", "DNS search order", (join ",",@{$AdapterInstance->{DNSServerSearchOrder}});
	    printf "%s: %s\n", "DNS domain", $AdapterInstance->{DNSDomain};
	    printf "%s: %s\n", "DNS domain suffix search order",  (join ",",@{$AdapterInstance->{DNSDomainSuffixSearchOrder}});
	    printf "%s: %s\n", "DNS enabled for WINS resolution", $AdapterInstance->{DNSEnabledForWINSResolution};
	}
	turn_on_errors();
}

#$class = get_class("Win32_LocalTime");
#print property_info($class->{Properties_}->{Day});
sub property_info{
   $prop = shift;
   $res  = $prop->{Name};
   $cimtype = $prop->{Qualifiers_}->Item("CIMType")->{Value};
   return $res."(".$cimtype.")";
 }

#print get_class_name(get_class("Win32_Directory"));
sub get_class_name{
	my $class = shift;
	return $class->{SystemProperties_}->{__CLASS}->{Value};
}

#print_class_attribute_qualifiers("Win32_LocalTime");
sub print_class_attribute_qualifiers{
	my $Class = shift; 
	$Class = get_class($Class) if (ref($Class) ne "Win32::OLE");

	print "De Property Qualifiers van alle attributen van de klasse ",get_class_name($Class),"\n\n";
	foreach my $prop (in $Class->{Properties_}){ #enkel de properties die specifiek zijn voor de klasse
	    print_attribute_qualifiers($prop);
	}
}

sub print_attribute_qualifiers{
	my $prop = shift;
	my $Qualifiers = $prop->{Qualifiers_};
	    printf "\n\n%s",$prop->{Name};
	    if ($Qualifiers->Item("CIMTYPE")){
	        printf " (%s <->%s = %s)",$prop->{CIMType},$Qualifiers->Item("CIMTYPE")->{Value},$cimtype{$prop->{CIMType}};     #de attribuutqualifiers bevat een duidelijke naam voor het type
	      }
	    printf "\n   Qualifiers: %s", join(" ",map {$_->{Name}} in $Qualifiers);
}

sub get_instances_from_class{
	my $Class = shift;
	$Class = get_class($Class) if (ref($Class) ne "Win32::OLE");
	return $Class->Instances_(wbemFlagUseAmendedQualifiers);
}

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
	$Win32::OLE::Warn = 0;
	my ($Object,$prop)=@_;
	return  $Object->{Qualifiers_}->Item($prop) && $Object->{Qualifiers_}->Item($prop)->{Value};
	$Win32::OLE::Warn = 3;
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
	print_attributes_from_instances(get_instances_from_class($class));
}

sub print_attributes_from_instances{
	my $instances = shift;
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
	
	print  get_class_name($class)," bevat ", $class->{Properties_}->{Count}," properties en ", 
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
	my %StartServiceReturnValues=make_method_qualifier_hash($Methods->Item("StartService"));
	my %StopServiceReturnValues=make_method_qualifier_hash($Methods->Item("StopService"));
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
	my %CreateReturnValues=make_method_qualifier_hash($methode);
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
	my %TerminateReturnValues=make_method_qualifier_hash($method);
	my $MethodInParameters =$method->{InParameters}; #Moeten echter niet ingevuld worden
	my $relpad="Win32_Process.Handle=\"$processHandle\"";
	my $object =  $WbemServices->get($relpad);
	print "$relpad wordt gestopt : ";
	my $MethodOutParameters=$object->ExecMethod_("Terminate",$MethodInParameters);
	$Return = $MethodOutParameters->{ReturnValue};
	print $TerminateReturnValues{$Return},"\n";
}
  
#maps the ValueMap to the Values for a given method
sub make_method_qualifier_hash{
	my $Method=shift;
	my %hash=();
	@hash{@{$Method->Qualifiers_(ValueMap)->{Value}}} = @{$Method->Qualifiers_(Values)->{Value}};
	return %hash;
}

#make hash from prop from class
sub make_property_hash{
	my ($prop,$Class)=@_;
	$Class = get_class($Class) if (ref($Class) ne "Win32::OLE");
	my $Qualifiers = $Class->Properties_($prop)->{Qualifiers_};
	my %hash=();
	@hash{@{$Qualifiers->Item("ValueMap")->{Value}}} = @{$Qualifiers->Item("Values")->{Value}};
	return %hash;
}


# EXCEL ###########################################################################################################


#initialisation
#@ARGV or die "give 1 argument: filename";
# my $filename = $ARGV[0];
# $filename = 'default.xls' if(!$filename);
# $fso = Win32::OLE->new("Scripting.FileSystemObject");

# $excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
# $excel->{DisplayAlerts}=0;
# $excel->{visible} = 1; 
#Main
#####################################
#$book = excel_open_file($filename);
# excel_print "\n Amount of cells:", get_amount_cells(), "\n";
# excel_print_empty_first_rows();
# excel_print_all_content_sheets();
# excel_print_specific_range("A1:D10");
# excel_print_cell(4,1);
# excel_print_range_between_cells(1,1,4,3);
#excel_update_cell(4,1,5);
#excel_update_range(1,1,8,8,"Excelsior");
#excel_multiplication_table(100,20);
#print "Press ENTER to exit"; <STDIN>;
#####################################
my ($excel, $fso);
sub excel_init{
	$excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
	$excel->{DisplayAlerts}=0;
	$excel->{visible} = 1; 
	$fso = Win32::OLE->new("Scripting.FileSystemObject");
}

sub excel_multiplication_table{
	my ($row_amount, $column_amount) = @_;
	my $book = excel_open_file("voud.xlsx");
	my $sheet = $book->Worksheets(1);
	$sheet->{name} = "Tables from 2 to 10";
	my $range = $sheet->Range($sheet->Cells(1,1),$sheet->Cells($row_amount, $column_amount));
	my $mat = $range->{Value};
	my $number=1;
	foreach my $row (@$mat) {
		$multiple=2;
		foreach (@$row){
			$_ = $number * $multiple; #writes to @$mat
			$multiple++;
		}
		$number++;
	}
	$range->{Value} = $mat;
	excel_add_table_lines($range);
}

sub excel_add_table_lines{
	my $range = shift;
	$range->Rows(1)->{font}->{bold} = 1;
	%constanten = %{Win32::OLE::Const->Load($excel)}; 
	$range->Borders($constanten{xlInsideVertical})->{LineStyle} = $constanten{xlContinuous}; 
	$range->Borders($constanten{xlEdgeRight})->{LineStyle} = $constanten{xlContinuous};
	$range->Borders($constanten{xlEdgeLeft})->{LineStyle} = $constanten{xlContinuous};
	$range->rows(1)->Borders($constanten{xlEdgeBottom})->{LineStyle} = $constanten{xlContinuous};
}

sub excel_update_cell{
	my ($row, $col, $value) = @_;
	my $sheet = $book->Worksheets(1);
	my $range = $sheet->Cells($row,$col);
	$range->{Value}=20;
	excel_save();
}

sub excel_update_range{
	my ($row1,$col1,$row2,$col2,$value) = @_;
	my $sheet = $book->Worksheets(1);
	for ($i=0;$i<$row2;$i++){#todo won't work if row1/col1 isn't 1,1
		for ($j=0; $j < $col2; $j++) {
			$mat->[$i][$j]=$value;
		}
	}
	$sheet->Range($sheet->Cells($row1,$col1),$sheet->Cells($row2,$col2))->{Value}=$mat;
	excel_save();
}

sub excel_save{
	$book->Save();
}

sub excel_get_start_next_row{
	my $last_cell = shift;
	$last_cell =~ s/\D//g; #digit only
	$last_cell++;
	return "A".$last_cell;	
}

sub excel_print_range_between_cells{
	my ($row1,$col1,$row2,$col2)= @_;
	my $sheet = $book->Worksheets(1);
	my $range = $sheet->Range($sheet->Cells($row1,$col1),$sheet->Cells($row2, $col2));
	excel_print_range($range);
}

sub excel_print_cell{
	my ($row, $col) = @_;
	my $sheet = $book->Worksheets(1);
	my $range = $sheet->Cells($row, $col);
	excel_print_range($range);
}

sub excel_print_specific_range{
	my $selected = shift;
	my $sheet = $book->Worksheets(1);
	my $range = $sheet->Range($selected);
	excel_print_range($range);
}

sub excel_print_content_all_sheets{
	foreach $nsheet (in $book->{Worksheets}){
	    print "\n$nsheet->{name}\n";
	    $range=$nsheet->{UsedRange};#filled cells
	    excel_print_range($range);
	}
}

sub excel_print_range{
	my $range = shift;
	my $mat = $range->{Value};
	if (ref $mat) { #multiple values
		print "\nmatrix with $range->{rows}->{Count} rows and $range->{columns}->{Count} columns\n";
		print join("  \t",@{$_}),"\n" foreach @{$mat};#double loop
	}
	else { #single value or empty
		($mat ? print "\n1 content : $mat\n": print "empty\n");
	}
	print "\n-----------------------------------------\n";
}

sub excel_print_empty_first_rows{
	foreach $nsheet (in $book->{Worksheets}){
		my $cell = $nsheet->Range("A1")->SpecialCells(xlCellTypeLastCell);
		my $range = $nsheet->Range("A1",$cell);
		printf "\n\t%-30s has %3d columns and %3d rows\n",$nsheet->{name},$range->{columns}->{count},$range->{rows}->{count} ;
	}
}

sub excel_get_amount_cells{
	my $sheet = $book->Worksheets(1);
	my $range_obj = $sheet->UsedRange();
	return $range_obj->{Count};
}

sub excel_get_amount_of_worksheets{
	print $book->Worksheets->{Count};
}

sub excel_open_file{
	my $filename = shift;
	my $book;
	if ($fso->FileExists($filename)) {
		my $path = $fso->GetAbsolutePathName($filename);
		$book=$excel->{Workbooks}->Open($path); 
	}
	else {
	    my $dir = $fso->GetAbsolutePathName("."); 
	    my $path = $dir."\\".$filename;            
	    $book = $excel->{Workbooks}->Add();   
	    $book->SaveAs($path);                 
	}
	return $book;
}