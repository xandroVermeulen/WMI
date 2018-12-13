#Todo vanaf reeks7.17
use Win32::OLE qw(EVENTS in);
use Win32::OLE::Const "Active DS Type Library";
$Win32::OLE::WARN = 1;
use Win32::OLE::Variant;
my $connection = make_connection();
my $RootObj = bind_object('RootDSE');
$RootObj->getInfo();
my $domein = $RootObj->{defaultNamingContext};

my $input;
if(@ARGV){
	handle_input2();
}
# while($input ne "q"){
# 	print "\nNext operation?:";
# 	my $input = <STDIN>;
# 	chomp $input;
# 	handle_input($input) if $input ne "q";
# }

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
  #$ADOconnection->Close();
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
	    PROPERTY_NOT_FOUND      => Win32::OLE::HRESULT(0x8000500D) );

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

use Math::BigInt;
sub convert_BigInt_string{
    my ($high,$low)=@_;
    my $HighPart = Math::BigInt->new($high);
    my $LowPart  = Math::BigInt->new($low);
    my $Radix    = Math::BigInt->new('0x100000000'); #dit is 2^32
    $LowPart+=$Radix if ($LowPart<0); #als unsigned int interperteren

    return ($HighPart * $Radix + $LowPart);
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
	    PROPERTY_NOT_FOUND      => Win32::OLE::HRESULT(0x8000500D) );

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
use Math::BigInt;
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

#todo: rewrite
sub get_moniker{
	my %monikers = (
		"loggedinhogent" => "LDAP://CN=Satan,OU=Domain Controllers,DC=iii,DC=hogent,DC=be",
		"loggedinhogent2" => "LDAP://iii.hogent.be/CN=Belial,OU=Domain Controllers,DC=iii,DC=hogent,DC=be",
		"homehogentip" => "LDAP://193.190.126.71/CN=Satan,OU=Domain Controllers,DC=iii,DC=hogent,DC=be",
		"homehogentdns" => "LDAP://satan.hogent.be/CN=Satan,OU=Domain Controllers,DC=iii,DC=hogent,DC=be",
		"loggedinugent" => "LDAP://CN=UGENTDC1,OU=Domain Controllers,DC=ugent,DC=be",
		"homeugentvpn" => "LDAP://ugentdc1.ugent.be:636/CN=UGENTDC2,OU=Domain Controllers,DC=ugent,DC=be",
		"xandro" => "CN=Xandro Vermeulen,OU=EM7INF,OU=Studenten,OU=iii,DC=iii,DC=hogent,DC=be"
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
