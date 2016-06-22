###	Sébastien Jardin
###	Resize massif PSV Photos
###	Release : 1.2
###	10/06/2015

#use strict;
#use warnings;

###Utilisation des modules
use Image::Magick;
use Spreadsheet::ParseExcel::Stream;
use Image::IPTCInfo;
use Array::Diff;
###

###Declarations des variables :
### Indiquez ici le dossier ou se trouve les photos à traiter, ne pas oublier les "\\"
###Exemple : my $targetPhotos = 'C:\\DOSSIER1\\SOUS-DOSIER2\\SOUS-DOSSIERxxx';
my $sourcePhotos = "..\\PHOTOS_IN";
###

###Indiquer ici la résolution souhaité
###Attention ne mettre que la largeur par exemple 800 pour un redimensionnement 800x532
###Exemple my $resolution = "800";
my $resolution = "800";
###

###Indiquer ici le fichier Excel
my $xls = Spreadsheet::ParseExcel::Stream->new('..\xls\result.xls');
###

###Variables Globales, ne pas toucher
my $i;
my $extension;
###

###Message d'accueil
print "\n \n"."Bonjour P.S.V \n \n";
print "\n"."Ce script va redimensionner les photos dans la resolution souhaite \n \n";
print "Le dossier source est celui-ci : ..\\PHOTOS_IN \n \n";
print "Le dossier cible est celui-ci : ..\\PHOTOS_OUT \n \n";
###

### Recuperation des photos a traiter
system("C:\\WINDOWS\\SYSTEM32\\cmd.exe /c dir /B $sourcePhotos > temp.liste");
###

###Gestion du dossier de sortie
print "Veuillez taper le nom du dossier d'epreuve : ";
$folder = <STDIN>;
chomp $folder;
print "Le dossier de sortie sera $folder \n";
###

### Indiquez ici le dossier ou se trouveront les photos traités, ne pas oublier les "\\"
###Exemple : my $targetPhotos = 'C:\\DOSSIER1\\SOUS-DOSIER2\\SOUS-DOSSIERxxx';
my $targetPhotos = "..\\PHOTOS_OUT\\$folder";
###

###Creation du repertoire mini dans target s'il n'existe pas
if (! -d $targetPhotos )
{
mkdir $targetPhotos;
}


### Indiquez ici le dossier ou se trouveront les mini photos traités, ne pas oublier les "\\"
###Exemple : my $targetPhotos = 'C:\\DOSSIER1\\SOUS-DOSIER2\\SOUS-DOSSIERxxx';
my $targetPhotosMinis = "$targetPhotos\\mini";
#print $targetPhotosMinis;
###

###Creation du repertoire mini dans target s'il n'existe pas
if (! -d $targetPhotosMinis )
{
mkdir $targetPhotosMinis;
}

### Indiquez ici le dossier ou se trouveront les photos non identifiées, ne pas oublier les "\\"
###Exemple : my $targetPhotosInconnus = 'C:\\DOSSIER1\\SOUS-DOSIER2\\SOUS-DOSSIERxxx';
my $targetPhotosInconnus = "$targetPhotos\\Inconnus";
#print $targetPhotosInconnus;
###

###Creation du repertoire mini dans target s'il n'existe pas
if (! -d $targetPhotosInconnus )
{
mkdir $targetPhotosInconnus;
}


### Indiquez ici le dossier ou se trouveront les photos non identifiées, ne pas oublier les "\\"
###Exemple : my $targetPhotosInconnusMinis = 'C:\\DOSSIER1\\SOUS-DOSIER2\\SOUS-DOSSIERxxx';
my $targetPhotosInconnusMinis = "$targetPhotos\\Inconnus\\mini";
#print $targetPhotosInconnusMinis;
###

###Creation du repertoire mini dans target s'il n'existe pas
if (! -d $targetPhotosInconnusMinis )
{
	mkdir $targetPhotosInconnusMinis;
}

###Affichage du nombre de photos à traiter
my $nb_lignes=0;
my $file = "temp.liste";
open(FILE, $file) or die "Can't open `$file': $!";
while (my $ligne = <FILE>) {
    $nb_lignes ++;
	chomp $ligne;
	#print "la ligne est :"."$ligne".": \n";
	$ligne =~ /(\w{4})(\w{4})(.*)/;
	$prefix = $1;
	$numphotos = $2;
	$extension = $3;
	#print "test du prefix depuis la liste : $prefix \n";
	#print "test du numéro de fichier \$2 : $2  \n";
	#print "test du nom de l'extension \$3 : $3 \n";
	###creation de la liste photos
	push @listephotos, "$numphotos";
	###
}
close (FILE);
print "\n Il y a $nb_lignes photo(s) a traiter. \n\n";
###


###Verification de l'extension
if ( $extension eq undef or $extension == "")
{
$extension = ".JPG";
}
###

###modification du prefix de la photo ### PHASE DE TEST###
$newprefix = "$folder"."_";
print "Le prefix : $prefix de la photo sera remplace par : $folder et devient : $newprefix \n";
###

###Time début traitement
my $time=time;
print "\n" ;
###Redimensionnement de la liste de photos
print "Debut du traitement photos : \n";
$i = 1;

### Re orientation des photos de manière automatique avec -auto-orient de mogrify
print "Re-orientation des photos pour les adapter au logiciel de vente \n";
system("C:\\ImageMagick-7.0.2-0-portable-Q16-x86\\mogrify.exe -auto-orient $sourcePhotos\\*.JPG");
###

###Declaration de la feuille 1 excel
my $sheet = $xls->sheet(1);
###

###Consommation de la 1ere ligne contenant la description de la colonne
$sheet->row;
###

###Seuil maximum du numéro de photo
$maxvalue = "9999";
#print "maxvalue : $maxvalue \n";
###


while ( my $row = $sheet->row )
{
	###liste @data contenant toutes les infos
	my @data = @$row;
	###
	
	###gestion des numéros de photos
	my $photos = @$row[1];
	#print "la liste \@photos est egale a : $photos \n";


	my @liste_photos_at_works = split /;/, $photos;

	foreach(@liste_photos_at_works)
	{
		#print "la liste a traiter est : $_ \n";
		$photos = $_;
		
		my ($num_photos_start, $num_photos_end) = split /[-,]/, $photos;
		#print "test split photos_start : $num_photos_start \n";
		#print "test split photos_end : $num_photos_end \n";
		
		
				
		###Verification de photos supplémentaires
		if (! defined $num_photos_end or $num_photos_end == "")
		{
			$num_photos_end = "$num_photos_start";
			#print" num_photos_end : $num_photos_end \n"
		}

		###Si photos sup à 9999
		if ($num_photos_start > $maxvalue or $num_photos_end > $maxvalue )
		{
			print "Attention il y a une erreur, il y a plus de 4 chiffres !\n";
			print "la liste de photos concerne est : $photos \n";
			print "Il faut corriger le fichier Excel !!! \n";
			exit ;
		}
		
		###Verification du matching avec le cavalier ou cheval
		if (@data[0] eq undef or @data[2] eq undef or @data[3] eq undef)
		{
			my $commentaire = "";
			push @liste_empty , $photos;
		}

		
		###Traitement photos : resize et commentaire
		foreach ("$num_photos_start".."$num_photos_end")
		{
			
			###Compensation de la notation excel et du numero photo
			#Application sur la photos en cours soit $_
			if ($_ =~ /^\d{1}$/)
			{
				$_ = "000"."$_";
				#print "photos traité est : $_ \n";
			}
			elsif ($_ =~ /^\d{2}$/)
			{
				$_ = "00"."$_";
				#print "photos traité est : $_ \n";
			}
			elsif ($_ =~ /^\d{3}$/)
			{
				$_ = "0"."$_";
				#print "photos traite est : $_ \n";
			}	
			
			
		
			my $source = "$sourcePhotos\\"."$prefix"."$_"."$extension";
			#print "La source est : $source \n";
			$sourceshort = "$prefix"."$_"."$extension";
			#print "la source short est :"."$sourceshort".": \n";
			my $dest = "$targetPhotos\\"."$newprefix"."$_"."$extension";
			#print "la destination est : $dest \n";
			my $destminis = "$targetPhotosMinis\\"."$newprefix"."$_"."$extension";
			#print "la destination est : $destminis \n";


			if ( -e $source)
			{
		
				#creation de la liste photos dans xls et noté par le photographe
				push @listephotosconnu, "$_";
				#
				#print "ceci est la variable \$_ : $_ \n";
				#print "ceci est la variable \$prefix\$_\$extension : "."$prefix"."$_"."$extension"." \n";
				print "Photo : $i sur $nb_lignes \n";

				###gestion du commentaire
				my $commentaire = "";
				#print "$commentaire \n\n";
				###	

				###Creation de l'objet pour la description source ou ORIGINAL
				#print $source;
				$info = new Image::IPTCInfo("$source");
				# Create object for file that may or may not have IPTC data.
				$info = create Image::IPTCInfo("$source");
				###
				###Insertion du commentaire dans l'objet
				$info->SetAttribute('caption/abstract', "$commentaire");
				#print $info;
				###Save du commentaire 
				$info->Save();
				$info->SaveAs($source);
			
				###gestion du commentaire
				my $commentaire = "@data[0] : @data[2] : @data[3]";
				#print "$commentaire \n\n";
				###	
			
				###Creation de l'objet pour la description source ou ORIGINAL
				#print $source;
				$info = new Image::IPTCInfo("$source");
				# Create object for file that may or may not have IPTC data.
				$info = create Image::IPTCInfo("$source");
				###
				###Insertion du commentaire dans l'objet
				$info->SetAttribute('caption/abstract', "$commentaire");
				#print $info;
				###Save du commentaire 
				$info->Save();
				$info->SaveAs($source);
				$info->SaveAs($dest);

				###Creation des miniatures
				my $image = Image::Magick->new;
				$image->Read("$source");
				$image->Resize(geometry => $resolution);
				$image->Write("$destminis");
				###			
				
				###Incrémentation de la photo
				$i ++;
			}
			else
			{
				##print "Attention photo supprime ou non note : $prefix"."$_"."$extension"." \n";
			}
		}
	}
}

print "\n";
print "**********************************\n";
print "Passage au traitement des Inconnus\n";
print "**********************************\n";
print "\n";

###Tries dans l'ordre des photos
###correction de bug sur le diff des 2 listes
###a ete contourne en triant dans le excel
@listephotosconnu = sort @listephotosconnu;
#print "la liste des photos connus est @listephotosconnu \n";
###

###Liste des photos non connus dans le Excel
my $diff = Array::Diff->diff( \@listephotos, \@listephotosconnu);
my @manquant = @{$diff->deleted()};
#print "la liste de photos manquantes est @manquant \n";
#my $diff_count = $diff->count;
my $size = @manquant;
###

###Mise a 0 du commentaire
my $commentaire = "00 : INCONNU : INCONNU";
#print "Le commentaire inconnu est : $commentaire \n" ;
###

if (@manquant eq undef or @manquant == "")
{
	##do nothing
}
else
{
	foreach (@manquant)
	{

		#print "la Photo $_ est inconnu.";
		#print "le \$_ est : $_ \n";

		my $source = "$sourcePhotos\\"."$prefix"."$_"."$extension";
		#print "La source est : "."$source"."\n";
		my $sourceshort = "$prefix"."$_"."$extension";
		#print "la source short est : $sourceshort \n";
		my $dest = "$targetPhotosInconnus\\"."$newprefix"."$_"."$extension";
		#print "la destination est : $dest \n";
		my $destminis = "$targetPhotosInconnusMinis\\"."$newprefix"."$_"."$extension";
		#print "la destination mini est : $destminis \n";

		#creation de la liste photos dans xls et noté par le photographe
		push @listephotosinconnus, "$prefix"."$_"."$extension";
		#
		#print "ceci est la variable \$numphotos : $numphotos \n";
		#print "ceci est la variable \$prefix\$numphotos\$extension : "."$prefix"."$numphotos"."$extension"." \n";
		print "Photo : $i sur $nb_lignes \n";


		###Creation de l'objet pour la description source ou ORIGINAL
		#print $source;
		$info = new Image::IPTCInfo("$source");
		# Create object for file that may or may not have IPTC data.
		$info = create Image::IPTCInfo("$source");
		###
		###Insertion du commentaire dans l'objet
		$info->SetAttribute('caption/abstract', "$commentaire");
		#print $info;
		###Save du commentaire 
		$info->Save();
		$info->SaveAs($source);
		$info->SaveAs($dest);
	
		###Creation des miniatures
		my $image = Image::Magick->new;
		$image->Read("$source");
		$image->Resize(geometry => $resolution);
		$image->Write("$destminis");
		###			

		###Incrémentation de la photo
		$i ++;

	}	
}
###


###Time fin traitement
$time=time-$time;
print "\nTemps d'execution en secondes : $time secondes \n\n";
###

###Liste photos inconus
print "Il y a $size photo(s) inconnus dans le Excel.\n";
print "la liste des photos inconnus et celle-ci :\n";
print "$_\n" for @listephotosinconnus;
###

###Liste photos ayant un problème
if (@liste_empty eq undef or @liste_empty == "")
{
	##do nothing
}
else
{
	print "Les photos suivantes sont mal identifiees : \n";
	print "$_\n" for @liste_empty;
}
###

###Message de fin
print "\n";
print "FIN du traitement, vous pouvez fermer cette fenetre";
###