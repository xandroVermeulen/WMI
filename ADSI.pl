#Todo vanaf reeks7.17
#Todo vanaf reeks8.14
#init
use Win32::OLE qw(EVENTS in);
use Win32::OLE::Const "Active DS Type Library";
use Math::BigInt;
$Win32::OLE::WARN = 1;
use Win32::OLE::Variant;
my $connection = make_connection();
my $RootObj = bind_object('RootDSE');
$RootObj->getInfo();
my $input;
my $domein = $RootObj->{defaultNamingContext};
my %E_ADS = (
    BAD_PATHNAME            => Win32::OLE::HRESULT(0x80005000),
    UNKNOWN_OBJECT          => Win32::OLE::HRESULT(0x80005004),
    PROPERTY_NOT_SET        => Win32::OLE::HRESULT(0x80005005),
    PROPERTY_INVALID        => Win32::OLE::HRESULT(0x80005007),
    BAD_PARAMETER           => Win32::OLE::HRESULT(0x80005008),
    OBJECT_UNBOUND          => Win32::OLE::HRESULT(0x80005009),
    PROPERTY_MODIFIED       => Win32::OLE::HRESULT(0x8000500B),
    OBJECT_EXISTS           => Win32::OLE::HRESULT(0x8000500E),
    SCHEMA_VIOLATION        => Win32::OLE::HRESULT(0x8000500F),
    COLUMN_NOT_SET          => Win32::OLE::HRESULT(0x80005010),
    ERRORSOCCURRED          => Win32::OLE::HRESULT(0x00005011),
    NOMORE_ROWS             => Win32::OLE::HRESULT(0x00005012),
    NOMORE_COLUMNS          => Win32::OLE::HRESULT(0x00005013),
    INVALID_FILTER          => Win32::OLE::HRESULT(0x80005014),
    INVALID_DOMAIN_OBJECT   => Win32::OLE::HRESULT(0x80005001),
    INVALID_USER_OBJECT     => Win32::OLE::HRESULT(0x80005002),
    INVALID_COMPUTER_OBJECT => Win32::OLE::HRESULT(0x80005003),
    PROPERTY_NOT_SUPPORTED  => Win32::OLE::HRESULT(0x80005006),
    PROPERTY_NOT_MODIFIED   => Win32::OLE::HRESULT(0x8000500A),
    CANT_CONVERT_DATATYPE   => Win32::OLE::HRESULT(0x8000500C),
    PROPERTY_NOT_FOUND      => Win32::OLE::HRESULT(0x8000500D)
);
#input
if(@ARGV){
	handle_input2();
}
my $multi_input_enabled = 0;
if($multi_input_enabled){
	while($input ne "q"){
		print "\nNext operation?:";
		my $input = <STDIN>;
		chomp $input;
		handle_input($input) if $input ne "q";
	}	
}

#only lab
sub delete_container{
	my $ou_naam=shift;
	my $ou=bind_object("ou=".$ou_naam.",ou=labo,".$domein);
	$ou->{Filter} = ["organizationalUnit"];  #enkel in de organizational unit wissen
	foreach (in $ou) {
	    delete_sub_containers($_);                      #wis alles in de container
	    print $_->{adspath}. " wissen ok (j=ja) ?\n"; #container zelf wissen
	    chomp($antw=<STDIN>);
	    if ($antw eq "j") {
	        $ou->delete ($_->{class},$_->{name});
	        if (Win32::OLE->LastError eq 0) {    
	        	print  $_->{adspath}," wordt gewist\n";
	    	} else {
	        	print Win32::OLE->LastError,"\n";
	    	}
	  	}
	}
}
#only lab
sub delete_sub_containers{
    my $cont=shift;
    foreach (in $cont){
        print $_->{adspath}. " wissen ok (j=ja) ?\n";
        chomp($antw=<STDIN>);
        if ($antw eq "j"){
            delete_sub_container($_);
            $cont->delete ($_->{class},$_->{name});
            if (Win32::OLE->LastError eq 0) {    
            	print  $_->{adspath}," wordt gewist\n";
            } else {
            	print Win32::OLE->LastError,"\n";
            }
     	}
   	}
}

#only lab
sub add_users_to_group{
	my $cont=bind_object("ou=...,ou=...,ou=labo,".$domein); #vul aan
	my $groep = bind_object("cn=...,ou=...,ou=...,ou=labo,".$domein);  #vul aan
    $cont->{filter}=[user];

    foreach (in $cont) { 
        push @leden ,$_->{distinguishedName};
    }

   $groep->PutEx(ADS_PROPERTY_UPDATE,"member",\@leden); #ADS_PROPERTY_UPDATE=2
   $groep->SetInfo();

   my $user=bind_object("cn=...,ou=...,ou=...,ou=labo,".domein);
   print join(",",@{get_attribute_value($user,"memberOf")});
}

#only lab
sub add_group_to_container{
	my $cont=bind_object("ou=...,ou=...,ou=labo,".$domein);
	my $groepnaam="groep";  # vul in naar keuze

	foreach (in $cont) { #make sure groepnaam is unique
        $_->GetInfoEx(["canonicalName"],0);
        $_->Get("canonicalName") =~ m/.*\/(.*)$/;
        lc($1) ne lc($groepnaam) or die "RDN moet uniek zijn !!"
	}

	my $command = ado_create_command();
	my $domeinobj = bind_object( $domein);
	my $sBase = $domeinobj->{adspath};
	#group, computer or person with SAM = groepnaam
	my $sFilter     = "(&(samAccountName=$groepnaam*)(|(objectcategory=group)(objectcategory=computer)(objectcategory=person)))";
	my $sAttributes = "samAccountName";
	ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes,"cn");

	my $ADOrecordset = ado_exec_query($command);
	my %lijst;
	until ( $ADOrecordset->{EOF} ) {
	    $lijst{$ADOrecordset->Fields("samAccountName")->{Value}}=1;
	    $ADOrecordset->MoveNext();
	}
	$ADOrecordset->Close();
	do {
	    $samnaam=sprintf("%s%02d",lc($groepnaam),++$tel);
	} while $lijst{$samnaam};

	my $groep=$cont->Create("group", "cn=$groepnaam");
	$groep->Put("samAccountName",$samnaam);
	$groep->SetInfo();
	print "toegevoegd met adspath: $groep->{adspath}\n"
	     unless (Win32::OLE->LastError());

	$groep->GetInfo();

	printf "%20s is ingesteld op %s\n",$_,join ("
	                                     ",@{get_attribute_value($groep,$_)})
	     foreach in bind_object($groep->{schema})->{MandatoryProperties};

}

sub print_groups_with_type{
	my @gtype=(ADS_GROUP_TYPE_GLOBAL_GROUP,ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP ,ADS_GROUP_TYPE_UNIVERSAL_GROUP);
	my @zoek=("Globale Beveiligingsgroep","Locale Beveiligingsgroep","Universele Beveiligingsgroep",
          "Globale Distributiegroep" ,"Locale Distributiegroep" ,"Universale Distributiegroep");
	print $_+1,": $zoek[$_]\n" foreach 0..$#zoek;

	do {
	  print "Kies een type groep: ";
	  chomp($nr=<STDIN>);
	} while ($nr>@zoek || $nr<1);
	$nr--;

	my $command = ado_create_command();
	my $domeinobj = bind_object($domein);
	my $sBase = $domeinobj->{adspath};
	my $sFilter     = "(&(&(objectcategory=group)(groupType:1.2.840.113556.1.4.803:=" . $gtype[$nr%3];
    $sFilter    .= $nr<3 ? ")(" : ")(!";
    $sFilter    .= "(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_SECURITY_ENABLED . "))))";
	my $sAttributes = "samAccountName,grouptype";
    ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "samAccountName");
	
	my $ADOrecordset = ado_exec_query($command);
	print "Overzicht $zoek[$nr]en:\n";
	until ( $ADOrecordset->{EOF} ) {
	    printf "\t%04b\t%s\n",$ADOrecordset->Fields("groupType")->{Value}%16    # 4 laagste bits
	                         ,$ADOrecordset->Fields("samAccountName")->{Value};
	    $ADOrecordset->MoveNext();
	}
	$ADOrecordset->Close();
}

sub print_hex_values_from_groups{
	my $command = ado_create_command();
	# alle groepen ophalen
	my $domeinobj = bind_object( $domein);
	my $sBase = $domeinobj->{adspath};
	my $sFilter     = "(objectcategory=group)";
	my $sAttributes = "cn,groupType";
	ado_create_query_from_command($command,$sBase, $sFilter, $sAttributes, "groupType");
	
	my $ADOrecordset = ado_exec_query($command);
	until ( $ADOrecordset->{EOF} ) {
	    printf "%04b\t%s\n",$ADOrecordset->Fields("groupType")->{Value}   # eerste bit staat op 1
	                       ,$ADOrecordset->Fields("cn")->{Value};
	    $ADOrecordset->MoveNext();
	}
	$ADOrecordset->Close();
}

#ToDo: oefening 8 geskipped.

#only lab
#update mail field based on user input for every object in the container
sub update_mail{
	my $cont=bind_object("ou=...,ou=...,ou=labo,".$RootObj->{defaultNamingContext}); # vul in
	$cont->{filter}=["user"];                   # enkel users in de container

	foreach (in $cont) {
	   print "mail(" . $_->Get("cn") . ") is ";
	   print $_->{mail} || "not set";
	   print  "\n\tgeef nieuw mail-adres: ";
	   chomp(my $nmail=<>);
	   $nmail ? $_->Put("mail",$nmail)  # ook mogelijk om het ADSI-attribuut te gebruiken:
	                                    # $_->{EmailAddress} = $nmail
	                                    # of met PutEx en update
	                                    # $_->PutEx(ADS_PROPERTY_UPDATE,"mail",[$nmail]);
	          : $_->PutEx(ADS_PROPERTY_CLEAR,"mail",[]); # geen equivalent mogelijk met het ADSI-attribuut
	   $_->SetInfo();
	}
}

#only in lab
sub add_user{

	my $cont=bind_object("ou=...,ou=...,ou=labo,".$RootObj->{defaultNamingContext}); # vul in
	my $usernaam="user_. . .";  # vul in max 20 tekens

	foreach (in $cont) {#make sure username doesn't exist yet in the container
	        $_->GetInfoEx(["canonicalName"],0);
	        $_->Get("canonicalName") =~ m/.*\/(.*)$/;
	        lc($1) ne lc($usernaam) or die "SPN moet uniek zijn !!"
	}
	#make sure one with the SAM doesn't already exist
	my $samnaam=$usernaam;
	my $command = ado_create_command();
	my $domeinobj = bind_object( $domein);
	my $sBase = $domeinobj->{adspath};
	my $sFilter     = "(&(samAccountName=$samnaam)(|(objectcategory=group)(objectcategory=computer)(objectcategory=person)))";
	my $sAttributes = "samAccountName";

	ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "cn");
	my $ADOrecordset = ado_exec_query($command);

	$ADOrecordset->{EOF} or die "Samnaam moet uniek zijn !!";
	$ADOrecordset->Close();
	#create the user
	my $user=$cont->Create("user", "cn=$usernaam");
	$user->Put("samAccountName",$samnaam);
	$user->SetInfo();
	print "toegevoegd met adspath: $user->{adspath}\n"
	     unless (Win32::OLE->LastError());

	$user->GetInfo();
	printf "%20s is ingesteld op %s\n",$_,join ("
	                                     ",@{get_attribute_value($user,$_)})
	     foreach in bind_object($user->{schema})->{MandatoryProperties};
}

sub print_all_SAM_names_and_longest{
	my $command = ado_create_command();
	my $domeinobj = bind_object($domein);

	my $sBase = $domeinobj->{adspath};
	my $sFilter     =  "(&(objectclass=user)(samAccountName=*))";
	my $sAttributes = "samAccountName,objectcategory";

    ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "cn");
	my $ADOrecordset = ado_exec_query($command);
	my $maxLength;
	until ( $ADOrecordset->{EOF} ) {
    	$samName=$ADOrecordset->Fields("samAccountName")->{Value};
    	$maxLength=length($samName) if (length($samName)>$maxLength);
    	print "$samName\n";
    	$ADOrecordset->MoveNext();
	}
	$ADOrecordset->Close();
	print "Maximale lengte is $maxLength\n";
}

#only in lab
sub add_container_to_my_container{
	my $cont=bind_object("ou= ... ,OU=Labo,".$RootObj->{defaultNamingContext});
	my %lijst; # lijst maken van alle subcontainers van cont
	foreach (in $cont) {
	        $_->GetInfoEx(["canonicalName"],0);
	        $_->Get("canonicalName") =~ m/.*\/(.*)$/;
	        $lijst{lc($1)}=undef;
	}

	my $ou_naam=shift;#naam ingeven voor nieuwe container
	while (exists $lijst{lc($ou_naam)} || !$ou_naam) {#tot iets nieuw opgegeven
	     print qq[canonicalName moet uniek zijn !\nde volgende namen mag je niet meer nemen in deze container: "]
	          ,join ('" "',keys %lijst),qq["\ngeef nieuwe naam:];
	     chomp($ou_naam=<STDIN>);
	}

	my $ou=$cont->Create("organizationalunit", "ou=$ou_naam");  #vergeet niet ou= toe te voegen in de tweede parameter.
	$ou->SetInfo();  #op het nieuwe object - niet op de container
	print "toegevoegd met verplichte properties ",join (", ",in bind_object($ou->{schema})->{MandatoryProperties})
	     unless (Win32::OLE->LastError());
}

#klasschema opvragen en dan bepaald verplicht attribuut kiezen voor te tonen(alle instanties).
sub print_mandatory_attribute_from_class{
	my $klassenaam = shift;
	my $klasse = bind_object( "schema/$klassenaam" );
	print "Verplicht attributen van $klassenaam:\n";
	!Win32::OLE->LastError()
		or die "je moet een ldapdisplayname van een klasse opgeven, vb container, organizationalUnit, ...\n";
	$klasse->{Class} eq "Class"
		or die "je moet een ldapdisplayname van een klasse opgeven, vb container, organizationalUnit, ...\n";

	my $tel=0;
	my @lijst=in $klasse->{MandatoryProperties};
	foreach (@lijst) {
	   $tel++;
	   print "$tel:\t$_\n";
	}

	my $nr;
	do { print "Kies een nummer <= $tel: ";
	     $nr=<STDIN>;
	   } until $nr>0 && $nr<=$tel;
	$nr--;

	my $Ldapdisplayname=$lijst[$nr];
	print "\nOverzicht van $Ldapdisplayname:\n";

	my $command = ado_create_command();
	my $domeinobj = bind_object($domein);
	my $sBase = $domeinobj->{adspath};
	my $sFilter     = "(objectclass=$klassenaam)";
	my $sAttributes = "distinguishedname";  #attribuut niet direct ophalen - kan geconstrueerd zijn ???
	ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "cn");
	
	my $ADOrecordset = ado_exec_query($command);
	until ($ADOrecordset->{EOF}) {
	    my $object=bind_object($ADOrecordset->Fields("distinguishedname")->{Value});
	    print join (";",@{valueattribuut($object,$Ldapdisplayname)}), " ($object->{name})\n"
	        if (uc($object->{class}) eq uc($klassenaam));
	    $ADOrecordset->MoveNext();
	}
	$ADOrecordset->Close();
}

#not possible at home
#maak studenten met zelfde description als mij lid van mijn groep
sub add_students_to_my_group{
	my $command = ado_create_command();

	my $groep=bind_object("cn=...,ou=...,ou=labo,". $RootObj->{defaultNamingContext}); #vul je eigen groep in
	$groep->GetInfoEx(["member"],0);
	my $descriptionvalue="*2/1*/*0"; #vervang dit door je eigen beschrijving

	my $domeinobj = bind_object($domein);
	my $sBase = $domeinobj->{adspath};
	my $sFilter     = "(&(description=" . $descriptionvalue . ")(objectclass=user)(objectcategory=person))";
	my $sAttributes = "distinguishedName";
	ado_create_query_from_command($command, $sBase,$sFilter,$sAttributes,"cn");

	my $ADOrecordset = ado_exec_query($command);
	my @lijst;
	until ( $ADOrecordset->{EOF} ) {
	    push @lijst,$ADOrecordset->Fields("distinguishedName")->{Value};
	    $ADOrecordset->MoveNext();
	}
	$ADOrecordset->Close();

	$groep->PutEx(ADS_PROPERTY_UPDATE,"member",\@lijst); #ADS_PROPERTY_UPDATE=2
	$groep->SetInfo() unless Win32::OLE->LastError();  
	print Win32::OLE->LastError(); #na setInfo wordt eventueel een fout gegeven
}


#return value from attribute in array
sub get_attribute_value {
    my ($object,$attribuut)=@_;
    my $attr_schema = bind_object( "schema/$attribuut" );
    my $tabel = $object->GetEx($attribuut);

    if (Win32::OLE->LastError() == Win32::OLE::HRESULT(0x8000500D)){
    	# maybe it wasn't cached
		$object->GetInfoEx([$attribuut], 0);
        $tabel = $object->GetEx($attribuut);
    }
    #if still not found
    return ["<niet ingesteld>"] if Win32::OLE->LastError() == Win32::OLE::HRESULT(0x8000500D);

    my $v=[];
    foreach ( in $tabel ) {
        if ( $attr_schema->{Syntax} eq "OctetString" ) {
            $waarde = sprintf ("%*v02X ","", $_) ;
        }
        elsif ( $attr_schema->{Syntax} eq "ObjectSecurityDescriptor" ) {
            $waarde = "eigenaar is ... " . $_->{owner};
        }
        elsif ( $attr_schema->{Syntax} eq "INTEGER8" ) {
            $waarde = convert_BigInt_string($_->{HighPart},$_->{LowPart});
        }
        else {
            $waarde = $_;
        }
        push @{$v},$waarde;
    }
    return $v;
}


################# reeks 8 ^  ##############

sub get_users_from_group{
	my $group = shift;
	$groep = "Labo" if (!$group);
	my $command = ado_create_command();
	my $domein  = bind_object( $RootObj->Get("defaultNamingContext") );
	my $sBase  = $domein->{adspath};
	
	my $sFilter = "(&(objectclass=group)(cn=$groep))"; #zoek de groep
	my $sAttributes = "distinguishedName"; 
	ado_create_query_from_command($command,$sBase, $sFilter, $sAttributes,"cn");
	my $ADOrecordset = ado_exec_query($command);
	($ADOrecordset->{RecordCount} ==1) or die "$groep bestaat niet\n";
	$distinguishedName =$ADOrecordset->Fields("distinguishedName")->{Value};
	print "$groep gevonden \nDistinguishedName = $distinguishedName \n";

	#alle leden ophalen met de matching rule
	my $sFilter = "(memberof:1.2.840.113556.1.4.1941:=$distinguishedName)";   #geen () toevoegen zoals in de documentatie van MSDN staat
	my $sAttributes = "cn,description,distinguishedName"; 
	ado_create_query_from_command($command,$sBase, $sFilter, $sAttributes,"cn");
	my $ADOrecordset = ado_exec_query($command);
	($ADOrecordset->{RecordCount} >=1) or die "$groep heeft geen users\n";
	print "\nbevat ", $ADOrecordset->{RecordCount}," users\n";
	$firstUser=$ADOrecordset->Fields("distinguishedName")->{Value}; #onthoud de eerste user
	until ( $ADOrecordset->{EOF} )  {
	    printf "%30s \t(%s)\n", $ADOrecordset->Fields("cn")->{Value},join("/",@{$ADOrecordset->Fields("description")->{Value}});
	   $ADOrecordset->MoveNext();
	}

	print "\nAlle groepen van $firstUser: \n";
	my $sFilter = "(member:1.2.840.113556.1.4.1941:=$firstUser)";   #geen () toevoegen zoals in de documentatie van MSDN staat
	my $sAttributes = "cn,description"; 
	ado_create_query_from_command($command,$sBase, $sFilter, $sAttributes,"cn");
	my $ADOrecordset = ado_exec_query($command);
	print $ADOrecordset->{RecordCount}," groepen:\n";
	until ( $ADOrecordset->{EOF} )  {
	    printf "%30s \t(%s)\n", $ADOrecordset->Fields("cn")->{Value},join("/",@{$ADOrecordset->Fields("description")->{Value}});
	    $ADOrecordset->MoveNext();
	}
}

sub print_search_and_systemflags{
	my $command = ado_create_command();
	my $schema  = bind_object( $RootObj->Get("schemaNamingContext") );
	my $sBase       =  $schema->{adspath};
	my $sFilter     = "(&(objectCategory=attributeSchema)"
	                . "(|(searchFlags:1.2.840.113556.1.4.803:=1)"
	                  . "(systemFlags:1.2.840.113556.1.4.804:=5)))";
	my $sAttributes = "cn,searchFlags,systemFlags";
    ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "cn");
	my $ADOrecordset = ado_exec_query($command);
	until ( $ADOrecordset->{EOF} ) {
	     my $prefix  = $ADOrecordset->Fields("searchFlags")->{Value} & 1 ? "I "  : "  ";
	        $prefix .= $ADOrecordset->Fields("systemFlags")->{Value} & 1 ? "NR " : "   ";
	        $prefix .= $ADOrecordset->Fields("systemFlags")->{Value} & 4 ? "C "  : "  ";
	        print $prefix , $ADOrecordset->Fields("cn")->{Value} , "\n";
	        $ADOrecordset->MoveNext();
	   }
	$ADOrecordset->Close();
}

sub print_attribute_schema_from_ldapname{

	my $ldapdisplayname = shift;
	my $command = ado_create_command();

	my $schema  = bind_object( $RootObj->Get("schemaNamingContext") );

	my $sBase  = $schema->{adspath};
	my $sFilter = "(ldapdisplayname=$ldapdisplayname)";
	my $sAttributes = "adspath"; #of name
    ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "cn");
    
    my $ADOrecordset = ado_exec_query($command);
	until ( $ADOrecordset->{EOF} )  {
	   print $ADOrecordset->Fields("adspath")->{Value},"\n";
	   $ADOrecordset->MoveNext();
	}
    $ADOrecordset->Close();

#Je kan dit ook als volgt ophalen, zonder LDAP-query (zie reeks 6 oefening 39)
#my $abstracteKlasse  = bind_object( "schema/$ldapdisplayname" );
#print "\nHet overeenkomstig reeel object heeft cn=",$abstracteKlasse->get("cn");
}

sub print_abstract_and_help_classes{
	my $command = ado_create_command();
	my $schema  = bind_object( $RootObj->Get("schemaNamingContext") );
	my $sBase  = $schema->{adspath};
	my $sFilter     = "(&(objectCategory=classSchema)(!(objectClassCategory=1)))";
	my $sAttributes = "cn,objectClassCategory";

	ado_create_query_from_command($command,$sBase, $sFilter, $sAttributes, "cn");
	my $ADOrecordset = ado_exec_query($command);

	print "Abstracte klassen:\n";
	until ($ADOrecordset->{EOF} ) {
	   print "\t" , $ADOrecordset->Fields("cn")->{Value} , "\n"  if( $ADOrecordset->Fields("objectClassCategory")->{Value} == 3 );
	   $ADOrecordset->MoveNext();
	}

	print "\nHulpklassen:\n";
	$ADOrecordset->MoveFirst();
	until ( $ADOrecordset->{EOF} )
	{
	   print "\t" , $ADOrecordset->Fields("cn")->{Value} , "\n" if( $ADOrecordset->Fields("objectClassCategory")->{Value} != 3 );        
	   $ADOrecordset->MoveNext();
	}

   $ADOrecordset->Close();
   $ADOconnection->Close();
}

sub find_group_name_from_student{
	my $command = ado_create_command();

	my $student ="....,OU=Studenten,ou=iii,".$RootObj->{defaultNamingContext}; #vul aan
	my $studentObject = bind_object($student);
	my $primaireGroep = $studentObject->get("primaryGroupID");

	my $domeinobj  = bind_object( $domein );
	my $sBase  = $domeinobj->{adspath};

	#alle groepen ophalen
	my $sFilter     = "(objectclass=group)";#(primaryGroupToken=$primaireGroep) lukt niet want geconstrueerd attribuut
	my $sAttributes =" primaryGrouptoken,cn";

	ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "cn");
	my $ADOrecordset = ado_exec_query($command);

	my $groepName="";
	while ($groepName eq "" && !$ADOrecordset->{EOF}) { 
	  my $primaryGroupToken = $ADOrecordset->Fields("primaryGrouptoken")->{Value};
	  if ($primaryGroupToken == $primaireGroep) {
	    $groepName=$ADOrecordset->Fields("cn")->{Value};
	  }
	  $ADOrecordset->MoveNext();
	}

	print $groepName;
    $ADOconnection->Close();
}

sub print_with_cn_filter{
	$command = ado_create_command();

	my $domeinobj  = bind_object( $domein );
	my $sBase  = $domeinobj->{adspath};
	my $sFilter     = "(&(objectCategory=computer)(cn=*A))";
	#my $sFilter     = "(&(objectCategory=computer)(name=*A))"; #lukt evengoed
	#my $sFilter     = "(&(objectCategory=computer)(canonicalname=*A))";    #geeft geen enkel resultaat
	my $sAttributes = "cn,canonicalName";

	ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "cn");

	my $ADOrecordset = ado_exec_query($command);
	until ( $ADOrecordset->{EOF} ) {
    	printf "%-20s %s %s\n",$ADOrecordset->Fields("cn")->{Value}
                          ,$ADOrecordset->Fields("canonicalName")->{Value}->[0];
    	$ADOrecordset->MoveNext();
    }
    $ADOrecordset->Close();
}

sub print_classes_with_attribute{
	my $ldapattribuut = shift;

	my $command = ado_create_command();
	my $sBase  = bind_object($domein)->{adspath};
	my $sFilter = "($ldapattribuut=*)";
	my $sAttributes = "$ldapattribuut,objectClass";

	ado_create_query_from_command($command, $sBase, $sFilter, $sAttributes, "cn");

	my $ADOrecordset = ado_exec_query($command);
	print "\n",$ADOrecordset->{RecordCount}," AD-objecten\n";
	my %classes;
	until ( $ADOrecordset->{EOF} )  {
	   $objectclass=$ADOrecordset->Fields("objectClass")->{Value};
	   $class=$objectclass->[-1]; #laatste element
	   $classes{$class}++;
	   $ADOrecordset->MoveNext();
	}
	$ADOrecordset->Close();
	foreach $cl (in keys %classes){
	   print $cl,"\n";
	}

#canonicalName is een geconstrueerd attribuut en mag niet in de filter worden opgenomen.
}

sub print_query_count{
	my $command = ado_create_command();

	my $container = bind_object($domein);
	my $sBase =$container->{adspath};

	my $sFilter     = ""; #mag leeg zijn
	#my $sFilter    = "(objectClass=user)";  #enkel users
	#my $sFilter    = "(objectClass=u*)";    #lukt niet met wildcards
	#my $sFilter    = "((distinguishedName=*) #lukt wel
	#my $sFilter    = "((distinguishedName=CN=*) #lukt niet meer
	#my $sFilter    = "(cn=*an*)";            #lukt wel met wildcards
	#my $sFilter    = "(cn=*)";               #attribuut cn moet ingesteld zijn

	my $sAttributes ="*";
	ado_create_query_from_command($command, $sBase,$sFilter, $sAttributes, "cn");
	my $ADOrecordset = ado_exec_query($command);

	print "\n",$ADOrecordset->{RecordCount}," AD-objecten \n";
    $ADOrecordset->Close();
}



sub find_students_with_postcode{
	my $postcode = shift;
	my $command = ado_create_command();

	my $distName="OU=studenten,OU=iii," . $domein;
	my $container = bind_object($distName);
	my $sFilter     = "(&(objectCategory=person)(objectclass=user)(postalCode=$postcode))"; 
	my $sBase =$container->{adspath};
	
	my $sAttributes = "cn,streetAddress,l";

	ado_create_query_from_command($command, $sBase,$sFilter, $sAttributes, "l");
	my $ADOrecordset = ado_exec_query($command);
	Win32::OLE->LastError() && die Win32::OLE->LastError();
	until ( $ADOrecordset->{EOF} )  {
      # printf "%-25s %s %s\n",$ADOrecordset->Fields("cn")->{Value}
      #                     ,$ADOrecordset->Fields("streetAddress")->{Value}
      #                     ,$ADOrecordset->Fields("l")->{Value};
        ado_show_record_fields($ADOrecordset);
    	$ADOrecordset->MoveNext();
	}  	
   	$ADOrecordset->Close();
}

sub find_container_with_name{
	my $name = shift;
	$name = "Stijn Van Hoecke" if (!name);
	my $ado_com = ado_create_command();

    my $domeinobj  = bind_object($domein);
    my $sBase  = $domeinobj->{adspath};
	my $sFilter     = "(name=$name)"; #maakt niet echt uit, maar geef wel een filter op
	my $sAttributes = "*";
	my $sSort      = "cn";

	ado_create_query_from_command($ado_com, $sBase,$sFilter, $sAttributes, $sSort);
	my $ADOrecordset = ado_exec_query($ado_com);
	ado_show_record_fields($ADOrecordset);
}

sub ado_print_recordset_info_from{
	#onderstaande oplossing resulteert in een fout - er ontbreekt een attribuut - zoek dit uit.
	my $ADOrecordset = shift;

	my $x=0;
    until ( $ADOrecordset->{EOF} ) {
      $x++;
      #todo fix for every object, not just printers
      print $x , "\t"
          , $ADOrecordset->Fields("printStaplingSupported")->{Value} , "\t"
          , $ADOrecordset->Fields("printRate")->{Value} , "\t"
          , $ADOrecordset->Fields("printMaxResolutionSupported")->{Value} , "\t"
          , $ADOrecordset->Fields("printerName")->{Value} , "\n";
      $ADOrecordset->MoveNext();
    }
    $ADOrecordset->Close();
}

sub ado_exec_query{
	my $ADOcommand = shift;
	my $ADOrecordset = $ADOcommand->Execute();
	Win32::OLE->LastError() && die (Win32::OLE->LastError());
	return $ADOrecordset;
}

sub ado_show_record_fields{
	my $ADOrecordset = shift;
	print $ADOrecordset->{RecordCount}," AD-objecten\n";
	print $ADOrecordset->{Fields}->{Count}," Ldap attributen opgehaald per AD-object\n";
	foreach my $field (in $ADOrecordset->{Fields}) {
	   print "\n$field->{name}($field->{type}):";
	   $waarde=$field->{value};
	   foreach (ref $waarde eq "ARRAY" ? @{$waarde} : $waarde) {
	      $field->{type} == 204
	         ? printf "\n\t%*v02X ","", $_
	         : print  "\n\t$_";
	   }
	}
}

sub ado_create_query_from_command{
	my ($ADOcommand,$sBase,$sFilter, $sAttributes, $sSort) = @_;

   #$sFilter     = "(&(objectCategory=printQueue)(printColor=TRUE)(printDuplexSupported=TRUE))";
   #$sAttributes = "printRate,printMaxXExtent,whenChanged,printStaplingSupported,objectClass,printMaxResolutionSupported,printername,adspath,objectGUID";

   my $sScope      = "subtree";
   $ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";
   $ADOcommand->{Properties}->{"Sort On"} = "$sSort";
}

my $ADOconnection;
sub ado_create_command{
   $ADOconnection = Win32::OLE->new("ADODB.Connection");
   $ADOconnection->{Provider} = "ADsDSOObject";
   if ( uc($ENV{USERDOMAIN}) ne "III") { #niet ingelogd op het III domein
       $ADOconnection->{Properties}->{"User ID"}          = "Xandro Vermeulen"; # vul in of zet in commentaar op school
       $ADOconnection->{Properties}->{"Password"}         = "Xandro Vermeulen"; # vul in of zet in commentaar op school
       $ADOconnection->{Properties}->{"Encrypt Password"} = True;
   }
   $ADOconnection->Open();  #!!

   my $ADOcommand = Win32::OLE->new("ADODB.Command");
   $ADOcommand->{ActiveConnection}      = $ADOconnection;        # verwijst naar het voorgaand object
   $ADOcommand->{Properties}->{"Page Size"} = 20;

	Win32::OLE->LastError()&& die (Win32::OLE->LastError());
	return $ADOcommand;
}

#algemene schema vragen, dan filteren op attributeschema
#dan voor elk object hiervan het abstract attribute shizzle opvragen
sub print_all_attributes_syntax{
	my $schema  = bind_object( $RootObj->Get("schemaNamingContext") );
	my %attributeSyntax = ();
	my %omSyntax = ();

	$schema->{Filter} = ["attributeSchema"];

	foreach my $attr_reeel ( in $schema ) {
	    my $attr_abstract = bind_object( "schema/" . $attr_reeel->{ldapDisplayName} );
	    $attributeSyntax{ $attr_abstract->{Syntax} } = $attr_reeel->{attributeSyntax};
	    $omSyntax{ $attr_abstract->{Syntax} }        = $attr_reeel->{omSyntax};
	}

	print "syntax van abstract object \t    syntax van reeel object\n";
	print "                           \tattributeSyntax \tomSyntax\n";
	print "-------------------------- \t---------------------------------\n";
	foreach my $i ( sort { substr( $attributeSyntax{$a}, 6 ) <=> substr( $attributeSyntax{$b}, 6 )
	                          || $omSyntax{$a} <=> $omSyntax{$b}
			   } keys %attributeSyntax ){

	    printf "%-24s\t%-18s\t%d\n", $i, $attributeSyntax{$i}, $omSyntax{$i};
	}
}

sub print_abstract_properties_admin_account{
	my $teller = 0;

	my $object = bind_object("CN=Administrator,CN=Users,".$RootObj->get(defaultNamingContext));
	my $abstracteKlasse = bind_object($object->{Schema});

	$object->GetInfoEx( $abstracteKlasse->{MandatoryProperties}, 0 );    # De Property Cache wordt ingevuld
	$object->GetInfoEx( $abstracteKlasse->{OptionalProperties} , 0 );

	foreach my $LDAPattribuut ( in $abstracteKlasse->{MandatoryProperties}, $abstracteKlasse->{OptionalProperties} ) 
	 {
	    $teller++;
	    my $abstractLdapAttribuut = bind_object( "schema/$LDAPattribuut" );
	    my $prefix =  "$teller: $LDAPattribuut  ($abstractLdapAttribuut->{Syntax})";
	    my $tabel = $object->GetEx($LDAPattribuut);

	    if ( Win32::OLE->LastError() == $E_ADS{PROPERTY_NOT_FOUND} )  {
	    	printlijn( \$prefix, "<niet ingesteld>" );
	    }
	    else {
	        foreach my $value ( @{$tabel} ) {
	            if ( $abstractLdapAttribuut->{Syntax} eq "OctetString" ) {
	                $waarde=sprintf ("%*v02X ","", $value) ;
	            }
	            elsif ( $abstractLdapAttribuut->{Syntax} eq "ObjectSecurityDescriptor" ) {
	                $waarde="eigenaar is ... $value->{owner}";
	            }
	            elsif ( $abstractLdapAttribuut->{Syntax} eq "INTEGER8" ){
	                $waarde=convert_BigInt_string($value->{HighPart},$value->{LowPart});
	            }
	            else {
	                $waarde=$value;
	            }
	            printlijn( \$prefix, $waarde );
			}
	    }
	}
}

sub printlijn {
    my ( $refprefix, $suffix ) = @_;
    printf "%-55s%s\n", ${$refprefix}, $suffix;
    ${$refprefix} = "";
}

sub print_adsi_properties_from_class{
	my $argument=shift;
	my $abstracteKlasse  = bind_object( "schema/$argument" );
	@attributen = qw(OID AuxDerivedFrom Abstract Auxiliary PossibleSuperiors MandatoryProperties
	                 OptionalProperties Container Containment);
	foreach my $prefix (@attributen){
	    my $attribuut = $abstracteKlasse->{$prefix}; #get the value 
	    printlijn( \$prefix, $_ ) foreach ref $attribuut eq "ARRAY" ? @{$attribuut} : $attribuut;
	}
}

sub print_abstract_schema{
	$abstractSchema = bind_object("schema");
	my %abstract;

	foreach (in $abstractSchema){
	    $abstract{$_->{class}}++;
	}

	while (($type,$aantal)=each %abstract){
	    print $type,"\t",$aantal,"\n";
	  }
}

sub print_schema_info{
	my $argument = shift; #either an attribute("Ldap-Display-Name") or class("user")
	my $object  = bind_object( "cn=" . $argument . "," . $RootObj->Get("schemaNamingContext") );

	if ( $object->{"Class"} eq "attributeSchema" ) {
	    @attributen = qw (cn distinguishedName canonicalName ldapDisplayName
	        attributeID attributeSyntax rangeLower rangeUpper
	        isSingleValued isMemberOfPartialAttributeSet
	        searchFlags  systemFlags);
	    }
	elsif ( $object->{"Class"} eq "classSchema" ) {
	    @attributen = qw(cn distinguishedName canonicalName ldapDisplayName
	        governsID subClassOf systemAuxiliaryClass AuxiliaryClass
	        objectClassCategory systemPossSuperiors possSuperiors
	        systemMustContain mustContain systemMayContain mayContain);
	    }
	else { die("cn=$argument niet gevonden in  reeel schema\n"); }

	$object->GetInfoEx([@attributen] , 0 );

	foreach my $attribuut (@attributen) {
	    my $prefix = $attribuut;
	    my $tabel  = $object->GetEx($attribuut);

	    if(Win32::OLE->LastError() == $E_ADS{PROPERTY_NOT_FOUND}){
	        printlijn( \$prefix, " < niet ingesteld > ");
	    }
	    else{
	        printlijn( \$prefix, $_ ) foreach @{$tabel};
	    }
	}
}

sub print_schema_classcount{
	my $reeelSchema  = bind_object( $RootObj->Get("schemaNamingContext") );
	my %reeel;
	foreach (in $reeelSchema){
	    $reeel{$_->{class}}++;
	}
	while (($type,$aantal)=each %reeel){
	    print $type,"\t",$aantal,"\n";
	}
}

sub print_attribute_from_user{
	my($wie,$property) = @_;

	#De basiscontainer is hard-gecodeerd.
	my $cont = bind_object("OU=EM7INF,OU=Studenten,OU=iii,".$RootObj->get(defaultNamingContext));
	my $user=$cont->getObject("user","cn=$wie");

	unless ($user) {
	    $cont = bind_object("OU=Docenten,OU=iii,".$RootObj->get(defaultNamingContext));
	    $user=$cont->getObject("user","cn=$wie");
	}

	unless($user ){
	    die( "$wie niet gevonden\n");
	}

	#drie regels die je moet aanpassen in volgende oefening
	$values=isset($user,$property); 
	defined($values) ||  die("$property is niet ingesteld\n");
	toon($property,$values);      
}

sub isset{
   my ($user,$property)=@_;
   $user->GetInfoEx([$property],0);# vraag expliciet om prop in de cache te plaatsen
   return $user->GetEx($property);# array referentie, undef indien niet ingevuld 
}

sub toon {
    my ($prop,$values)=@_;
    $prefix=$prop;
    printlijn( \$prefix, $_ ) foreach @{$values};
}

sub print_property_cache{
	my $administrator = bind_object("CN=Administrator,CN=Users,".$RootObj->get(defaultNamingContext));
	$administrator->GetInfo();

	print "Aantal attributen in de Property Cache: $administrator->{PropertyCount}\n";

	for ( my $i = 0 ; $i < $administrator->{PropertyCount}; $i++ ){

	    my $attribuut = $administrator->Next();
	#   my $attribuut = $administrator->Item($i); #alternatief
	    my $prefix    = $attribuut->{Name} . " (" . $attribuut->{ADsType} . ")";
	    foreach my $propValue ( @{$attribuut->{Values}} )  {
	        if ( $attribuut->{ADsType} == ADSTYPE_NT_SECURITY_DESCRIPTOR) {
	             $sec_object=$propValue->GetObjectProperty($attribuut->{ADsType});
	             $suffix = "eigenaar is ..." . $sec_object->{owner};
	        }
	        elsif ( $attribuut->{ADsType} == ADSTYPE_OCTET_STRING) {
	            $inhoud=$propValue->GetObjectProperty($attribuut->{ADsType});
	            $suffix = sprintf "%*v02X ","",$inhoud;
	        }
	        elsif ( $attribuut->{ADsType} == ADSTYPE_LARGE_INTEGER ) {
	            $inhoud=$propValue->GetObjectProperty($attribuut->{ADsType});
	            $suffix = convert_BigInt_string($inhoud->{HighPart},$inhoud->{LowPart});
	        }
	        else  {
	            $suffix = $propValue->GetObjectProperty($attribuut->{ADsType});
	            }
	        printlijn( \$prefix, $suffix );
	   }
	}


}
#een groot geheel getal wordt teruggegeven als twee gehele getallen
#vb  29868835, -1066931206
# Het groot geheel getal dat hierbij hoort moet je berekenen als volgt :
# 29868835 . 2^32 + (2^32 - 1066931206 ) = 128285672722656250
# Onderstaande functie berekent deze waarde met behulp van de module Math::BigInt.
# Een bijkomend probleem is dat je deze waarde enkel met print juist uitschrijft,
# gebruik je printf dan moet je %s en niet %g gebruiken, anders krijg je een "afgeronde" waarde.

sub convert_BigInt_string{
    my ($high,$low)=@_;
    my $HighPart = Math::BigInt->new($high);
    my $LowPart = Math::BigInt->new($low);
    my $Radix = Math::BigInt->new('0x100000000'); #dit is 2^32
    $LowPart+=$Radix if ($LowPart<0); #als unsigned int interperteren

    return ($HighPart * $Radix + $LowPart);
}

sub printlijn {
    my ( $refprefix, $suffix ) = @_;
    printf "%-35s%s\n", ${$refprefix}, $suffix;
    ${$refprefix} = "";
}

sub print_subcontainer_child_estimation{
	my $cont = bind_object("OU=Studenten,OU=iii,".$RootObj->get(defaultNamingContext));

	print "Groepen:\n";
	$cont->{Filter} = ["organizationalUnit"];
	foreach my $subcont (in $cont) {
	    # het betreft een geconstrueerd LDAP-attribuut !
	    $subcont->GetInfoEx(["ou","msDS-Approx-Immed-Subordinates"],0);
	    $waarde=$subcont->Get("msDS-Approx-Immed-Subordinates");
	    printf "% 7s: %d\n" ,$subcont->{ou}, $waarde;
	}
}

sub print_main_partitions{
	print "Domeingegevens:        $RootObj->{defaultNamingContext}\n";
	print "Configuratie gegevens: $RootObj->{configurationNamingContext}\n";
	print "Schema:                $RootObj->{schemaNamingContext}\n";
}

sub print_all_partitions{
	print join("\n",@{$RootObj->{"namingContexts"}});
}

sub show_users_type{
	my $class=shift;
	my $Users = bind_object("CN=Users,$domein");
	$Users->{Filter} = [$class];
	print "AD-objecten van type $class:\n";
	print "$_->{adspath}\n" foreach in $Users;
}
sub initialize_user_child{
	print "\nAdministrator:\n";
	my $a = $Users->GetObject("user","CN=Administrator");
	print $a->{ADsPath};
}

sub print_system_container{
	my $systeemContainer=bind_object("cn=system,$domein");
	foreach (in $systeemContainer){
    	printf "%35s: %s\n", $_->{Name},$_->{class};
	}
}

#	"geef computerklas 215 219 223 225";
sub get_computers_in_room{
	my $lokaal = shift;
	my $lokaalContainer=bind_object("OU=$lokaal,OU=PC's,OU=iii,$domein");
	print $lokaalContainer->{adspath},"\n";
	foreach (in $lokaalContainer){
	    print $_->{cn},"\n";
	}
}

#analoog    dsquery computer "ou=Domain Controllers,dc=iii,dc=hogent,dc=be" 
sub get_objects_in{
	my $containername = shift;
	$containername = "ou=Domain Controllers" if (!$containername);
	my $container=bind_object("$containername,$domein");
	foreach (in $container){
	    print $_->{cn},"\n";
	}
}

sub print_directory_service_info{
	my $DomeinDN = GetDomeinDN();
	my $o = bind_object("CN=Administrator,CN=Users,$DomeinDN");
	print "\n";
	print "RDN:                   $o->{Name}\n";
	print "klasse:                $o->{Class}\n";
	print "objectGUID:            $o->{GUID}\n";
	print "ADsPath:               $o->{ADsPath}\n";
	print "ADsPath Parent:        $o->{Parent}\n";
	print "ADsPath schema klasse: $o->{Schema}\n";
}

sub GetDomeinDN {
    print "Server DNS:               $RootObj->{dnsHostName}\n";
    print "       SPN:               $RootObj->{ldapServiceName}\n";
    print "       Datum & tijd:      $RootObj->{currentTime}\n";
    print "       Global Catalog ?   $RootObj->{isGlobalCatalogReady}\n";
    print "       gesynchronizeerd ? $RootObj->{isSynchronized}\n";
    print "       DN:                $RootObj->{serverName}\n";

    print "Domeingegevens:           $RootObj->{defaultNamingContext}\n";
    print "Configuratiegegevens:     $RootObj->{configurationNamingContext}\n";
    print "Schema:                   $RootObj->{schemaNamingContext}\n";

    print "Functioneel niveau \n";
    print "    Forest:               $RootObj->{forestFunctionality}\n";
    print "    Domein:               $RootObj->{domainFunctionality}\n";

    return $RootObj->{defaultNamingContext};
}

sub print_object{
	my $Adspath = shift;
	$AdsPath = "OU=EM7INF,OU=Studenten,OU=iii,DC=iii,DC=hogent,DC=be" if (!$Adspath);
	my $obj     = bind_object($AdsPath);
	print "--------------------ADSI-------------------------------\n";
	print  "Adspath (ADSI) = " ;
	printcontent ($obj->{ADsPath});

	print  "class (ADSI)   = ";
	printcontent ($obj->{class});

	print  "GUID (ADSI)    = ";
	printcontent ($obj->{GUID});

	print "--------------------LDAP-------------------------------\n";
	print  "distinguishedName (LDAP) = " ;
	printcontent ($obj->{distinguishedName});

	print  "objectclass (LDAP)       = ";
	printcontent ($obj->{objectclass});

	print  "objectGUID (LDAP)        = ";
	printcontent (sprintf ("%*v02X ","",$obj->{objectGUID}));

}

sub printcontent{
   my $inhoud=shift;
   if (ref $inhoud) {
       print "Array met " , scalar @{$inhoud} , " elementen :\n\t" ;
       print join("\n\t", @{$inhoud});
   }
   else {
       print "$inhoud";
   }
   print "\n";

}

sub get_user{
	return bind_object(get_moniker("Xandro"));
}

sub print_user{
	my $user = get_user();
	printf "%-20s : %s\n", $_ , $user->{$_} foreach qw (mail givenName sn displayName homeDirectory scriptPath profilePath logonHours userWorkstations);
}
#sub test{print_ADSI_attributes(get_user());}
sub print_ADSI_attributes{
	my $obj = shift;
	my @attr = ("adspath", "class", "GUID", "name", "parent", "Schema");
	foreach (@attr){
		printf "%20s : %s\n", $_, $obj->{$_}; 
	}
}

sub handle_input2{
	$method = shift @ARGV;
	&$method(@ARGV);
}

sub handle_input{
	my @data = split / /, shift;
	my $method = shift @data;
	&$method(@data);
}

sub get_moniker{
	my %monikers = (
		"loggedinhogent" => "LDAP://CN=Satan,OU=Domain Controllers,".$domein,
		"loggedinhogent2" => "LDAP://iii.hogent.be/CN=Belial,OU=Domain Controllers,".$domein,
		"homehogentip" => "LDAP://193.190.126.71/CN=Satan,OU=Domain Controllers,".$domein,
		"homehogentdns" => "LDAP://satan.hogent.be/CN=Satan,OU=Domain Controllers,".$domein,
		"loggedinugent" => "LDAP://CN=UGENTDC1,OU=Domain Controllers,DC=ugent,DC=be",
		"homeugentvpn" => "LDAP://ugentdc1.ugent.be:636/CN=UGENTDC2,OU=Domain Controllers,DC=ugent,DC=be",
		"xandro" => "CN=Xandro Vermeulen,OU=EM7INF,OU=Studenten,OU=iii,".$domein
		);	
	return $monikers{lc(shift)}
}

sub make_connection{
	my $obj = bind_object(get_moniker("homehogentip"));
	print (Win32::OLE->LastError()?"not good":"Connected:"),"\n";
	return $obj;
}

sub bind_object{
	my $param = shift;
	my $monik;
	if(uc($ENV{USERDOMAIN}) eq "III"){
		$monik = (uc(substr($param, 0,7)) eq "LDAP://" ? "" : "LDAP://").$param;
		return (Win32::OLE->GetObject($monik));
	} else {
		my $dso = Win32::OLE->GetObject("LDAP:");
		my $ip = "193.190.126.71";
		$monik = (uc(substr($param, 0,7)) eq "LDAP://" ? "" : "LDAP://$ip/").$param;
		my $login = "Xandro Vermeulen";
		Win32::OLE->LastError() and die Win32::OLE->LastError();
		return ($dso->OpenDSObject($monik,$login,$login,ADS_SECURE_AUTHENTICATION));
	}
}

sub show_constants{
	my %const = %{Win32::OLE::Const->Load("Active DS Type Library")};
	while (($key, $value) = each (%const)){
		printf "%s: %s\n", $key, $value;
	}
}
