// Antes de usar estos macros, se necesita descargar Read and Write Excel
// Este plugin permite que todos los resultados que se vayan recatando se vayan directo a un excel
// Las instrucciones para descargarlo estan aqui: https://imagej.net/plugins/read-and-write-excel
// (Al final de la pagina)

macro "Get nucleos [n]" {

	dir=getDirectory("Choose a Directory");
	lst=getFileList(dir);
	run("Set Measurements...", "area mean integrated redirect=None decimal=5");
	
	for ( i=0; i<lst.length; i++ ) {
		if(endsWith(dir+lst[i], ".lsm")){
   		open(dir+lst[i]);
		run("Split Channels");
		close("C1-" + lst[i]);
		close("C2-" + lst[i]);
		selectImage("C4-" + lst[i]);
		run("Duplicate...", " ");
		run("Tile");
		selectImage("C4-" + lst[i]);
		rename("DAPI-" + lst[i]);
		selectWindow("DAPI-" + lst[i]);
		run("Median...", "radius=3");
		setAutoThreshold("Moments dark");
		waitForUser("Seleccionar los nucleos ajustando threshold, contraste y utilizando el pincel. O como quieras la verdad. Lo importante es que todo este seleccionado.");
		setOption("BlackBackground", false);
		run("Convert to Mask");
		run("Fill Holes");
		run("Analyze Particles...", "size=10-Infinity add include");
		waitForUser("Todo OK?", "Si no te gusta algunos de los ROI, puedes eliminarlo en el roi manager");
		roiManager("Save", dir + lst[i] + "RoiSet-DAPI.zip");
		roiManager("measure");
		path = dir + "Nucleos.xlsx";
		run("Read and Write Excel", "dataset_label=[Atributos de los nucleos medidos] stack_results file=["+path+"]" );
		roiManager("Deselect");
		roiManager("Delete");
		selectWindow("ROI Manager");
		wait(100);
		run("Close");
		wait(100);
		selectWindow("Results");
		wait(100);
		run("Close");
		close("*");
		}
	}
	
}

macro "Get cell [c]" {
	dir=getDirectory("Choose a Directory");
	lst=getFileList(dir);
	run("Set Measurements...", "area mean integrated redirect=None decimal=5");
	
	for ( i=0; i<lst.length; i++ ) {
		if(endsWith(dir+lst[i], ".lsm")){
			open(dir+lst[i]);
			run("Split Channels");
			close("C1-" + lst[i]);
			close("C3-" + lst[i]);
			selectImage("C4-" + lst[i]);
			run("Enhance Contrast", "saturated=0.35");
			rename("DAPI-" + lst[i]);
			selectImage("C2-" + lst[i]);
			rename("Red-" + lst[i]);
			selectWindow("Red-" + lst[i]);
			run("Merge Channels...", "c1=[Red-" + lst[i] + "] c3=[DAPI-" + lst[i] + "]");
			run("Tile");
			waitForUser("Seleccionar la celula utilizando el pincel en verde. O como quieras la verdad. Lo importante es que todo este seleccionado.");
			run("Split Channels");
			selectImage("RGB (green)");
			run("Convert to Mask");
			run("Fill Holes");
			run("Analyze Particles...", "size=40-Infinity add include");
			waitForUser("OK?");
			roiManager("Save", dir + lst[i] + "RoiSet-Cell.zip");
			roiManager("Deselect");
			roiManager("Delete");
			run("Close All");
		}
	}

}


macro "Get localizacion Red vs DAPI [l]" {
	
	
	dir=getDirectory("Choose a Directory");
	lst=getFileList(dir);
	run("Set Measurements...", "area mean integrated redirect=None decimal=5");
	
	for ( i=0; i<lst.length; i++ ) {
		if(endsWith(dir+lst[i], ".lsm")){
			open(dir+lst[i]);
			run("Split Channels");
			close("C1-" + lst[i]);
			close("C4-" + lst[i]);
			close("C3-" + lst[i]);
			selectImage("C2-" + lst[i]);
			rename("Red-" + lst[i]);
			roiManager("Open", dir + lst[i] + "RoiSet-Cell.zip");
			roiManager("Open", dir + lst[i] + "RoiSet-DAPI.zip");
			roiManager("Show All with labels");
			c = roiManager("count");
			for (j = 0; j < c; j++){
				for (k = j + 1; k < c; k++){
					roiManager("Select", newArray(j,k));
					roiManager("AND");
					if (selectionType()>-1) {
						if (getBoolean("¿Correcto?")) {
							roiManager("Select", newArray(j,k));
							roiManager("measure");
							roiManager("XOR");
							run("Measure");
						}
					}
				}
			}
			path = dir + "Red.xlsx";
			run("Read and Write Excel", "file=["+path+"]" );
			roiManager("Deselect");
			roiManager("Delete");
			selectWindow("ROI Manager");
			run("Close");
			selectWindow("Results");
			run("Close");
			close("*");
			run("Close All");
			wait(100);
		}
	}
	
}


macro "Get Lamin B1 [b]" {
	
	
	dir=getDirectory("Choose a Directory");
	lst=getFileList(dir);
	run("Set Measurements...", "area mean integrated redirect=None decimal=5");
	
	for ( i=0; i<lst.length; i++ ) {
		if(endsWith(dir+lst[i], ".lsm")){
			open(dir+lst[i]);
			run("Split Channels");
			close("C1-" + lst[i]);
			close("C4-" + lst[i]);
			close("C2-" + lst[i]);
			selectImage("C3-" + lst[i]);
			if (getBoolean("¿Es una imagen de Lamin B1?")) {
				roiManager("Open", dir + lst[i] + "RoiSet-DAPI.zip");
				c = roiManager("count");
				while (c > 0) {
					roiManager("Select", 0);
					run("Enlarge...", "enlarge=-1.5");
					run("Make Band...", "band=2");
					if (getBoolean("¿Satisfactorio?")){
						roiManager("Add");
					}
					roiManager("Select", 0);
					roiManager("Delete");
					c = c - 1;
				}
				roiManager("measure");
				roiManager("Delete");
			}
			close("C3-" + lst[i]);		
		}
	}
	path = dir + "LaminB1.xlsx";
	run("Read and Write Excel", "file=["+path+"]" );
	selectWindow("ROI Manager");
	run("Close");
	selectWindow("Results");
	run("Close");
	close("*");
	run("Close All");
	wait(100);		
}





macro "Get verde [v]" {

	dir=getDirectory("Choose a Directory");
	lst=getFileList(dir);
	run("Set Measurements...", "area mean integrated redirect=None decimal=5");
	
	for ( i=0; i<lst.length; i++ ) {
		if(endsWith(dir+lst[i], ".lsm")){
   		open(dir+lst[i]);
		run("Split Channels");
		close("C1-" + lst[i]);
		close("C2-" + lst[i]);
		close("C4-" + lst[i]);
		roiManager("Open", dir+lst[i] + "RoiSet-DAPI.zip");
		roiManager("measure");
		path = dir + "Nuclear_Green_Signal.xlsx";
		run("Read and Write Excel", "file=["+path+"]" );
		roiManager("Deselect");
		roiManager("Delete");
		selectWindow("ROI Manager");
		wait(100);
		run("Close");
		wait(100);
		selectWindow("Results");
		wait(100);
		run("Close");
		close("*");
		}
	}
	
}