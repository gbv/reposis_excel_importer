package exelparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Properties;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.Namespace;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;
import org.junit.Test;
import org.mycore.common.MCRConstants;
import org.mycore.common.MCRTestCase;
import org.mycore.common.events.MCRStartupHandler;
import org.mycore.mods.MCRMODSWrapper;

public class PoiConvertXlsxToMods extends MCRTestCase {

	private final static Logger LOGGER = LogManager.getLogger(MCRStartupHandler.class);

	@Override protected Map<String, String> getTestProperties() {
		Map<String, String> props = super.getTestProperties();
		props.put("MCR.MODS.NewObjectType","mods");
		props.put("MCR.Metadata.Type.mods", "true");
		props.put("MCR.MODS.Types", "mods");
		return props;
	}

    public class Person {
        private String sName;
        private String fName;
        private String dnbID;

        public Person(String sName, String fName, String dnbID) {
            this.sName = sName;
            this.fName = fName;
            this.dnbID = dnbID;
        }

        public String getsName() {
            return sName;
        }

        public String getfName() {
            return fName;
        }

        public String getDnbID() {
            return dnbID;
        }
    }

    public class RolePersonTuple {
	    private String role;
	    private Person person;

        public RolePersonTuple(String role, Person person) {
            this.role = role;
            this.person = person;
        }

        public String getRole() {
            return role;
        }

        public Person getPerson() {
            return person;
        }
    }

	@Test
	public void test() throws IOException {
        Properties properties = new Properties();
        InputStream input = new FileInputStream("config.properties");
        properties.load(input);

        try (XSSFWorkbook ggArchivDaten = new XSSFWorkbook(new FileInputStream(
            new File(properties.getProperty("sourcePath") + "GrassArchivDaten.xlsx")));
            XSSFWorkbook ggBearbeiter = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "Bearbeiter.xlsx")));
            XSSFWorkbook ggSach = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "SWSach.xlsx")));
            XSSFWorkbook ggWerke = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "SWWerke.xlsx")));
            XSSFWorkbook ggGeo = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "GeoDaten.xlsx")));
            XSSFWorkbook ggZeit = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "SWZeit.xlsx")));
            XSSFWorkbook ggInstitutionen = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "Rechte.xlsx")));
            XSSFWorkbook ggPersonenAlle = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "PersonenteilnahmeAlle.xlsx")));
            XSSFWorkbook ggPersonenNeu = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "PersonenNeu.xlsx")));
            XSSFWorkbook ggPersonenRollen = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "Personenrollen.xlsx")));
            XSSFWorkbook ggSWPersonen = new XSSFWorkbook(new FileInputStream(
                new File(properties.getProperty("sourcePath") + "GrassArchivDaten_SWPerson.xlsx")));

        ) {

            XSSFSheet archivDaten = ggArchivDaten.getSheetAt(0);
            XSSFSheet editor = ggBearbeiter.getSheetAt(0);
            XSSFSheet swSach = ggSach.getSheetAt(0);
            XSSFSheet swWerke = ggWerke.getSheetAt(0);
            XSSFSheet geoDaten = ggGeo.getSheetAt(0);
            XSSFSheet zeitDaten = ggZeit.getSheetAt(0);
            XSSFSheet institutionen = ggInstitutionen.getSheetAt(0);
            XSSFSheet personenAlle = ggPersonenAlle.getSheetAt(0);
            XSSFSheet personenNeu = ggPersonenNeu.getSheetAt(0);
            XSSFSheet personenRollen = ggPersonenRollen.getSheetAt(0);
            XSSFSheet swPersonen = ggSWPersonen.getSheetAt(0);

            MCRMODSWrapper mcrmodsWrapper = new MCRMODSWrapper();
            mcrmodsWrapper.getMCRObject().setVersion("test"); //set MyCoRe-Version
            mcrmodsWrapper.setMODS(new Element("mods", MCRConstants.MODS_NAMESPACE));

            Map<Integer, String> tableHeaderMap = new HashMap<>();
            XSSFRow rowHeader = archivDaten.getRow(0);
            for (Cell c : rowHeader) {
                tableHeaderMap.put(c.getColumnIndex(), c.getStringCellValue());
            }

            for (int i = 2; i < archivDaten.getPhysicalNumberOfRows(); i++) {
                String formatted = String.format("%08d", i); //creates an ongoing id, like 00000001 ... 00000122 ...
                String savePath = properties.getProperty("savePath") + formatted + ".xml";
                XSSFRow row = archivDaten.getRow(i);
                for (int columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++) {
                    Cell cell = row.getCell(columnIndex);
                    String columnName = tableHeaderMap.get(columnIndex);
                    if (cell != null) {
                        switch (columnName) {
                            case "GGID":
                                cell.setCellType(CellType.STRING);
                                modsBuildHeader(mcrmodsWrapper, i);
                                //modsLocation(mcrmodsWrapper, id);
                                break;
                            case "Mediensignatur":
                                if (cell.getStringCellValue().equals("VID")) {
                                    modsGenre(mcrmodsWrapper, "video", "moving image");
                                } else {
                                    modsGenre(mcrmodsWrapper, "audio", "sound recording");
                                }
                                break;
                            case "Titel":
                                if (row.getCell(columnIndex + 1) != null) {
                                    Cell cellSubTitle = row.getCell(columnIndex + 1);
                                    modsTitle(mcrmodsWrapper, cell, cellSubTitle, "alternative");
                                } else {
                                    modsTitle(mcrmodsWrapper, cell, null, "alternative");
                                }
                                break;
                            case "Kontext":
                                modsNote(mcrmodsWrapper, cell, "context");
                                break;
                            case "Abstract":
                                modsAbstract(mcrmodsWrapper, cell);
                                break;
                            case "Bearbeiter":
                                String editorName = checkEditorName(editor, cell.getStringCellValue());
                                modsNameEditor(mcrmodsWrapper, cell, editorName);
                                break;
                            case "Anmerkungen":
                                modsNote(mcrmodsWrapper, cell, "content");
                                break;
                            case "Tonträger":
                                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    modsClassification(mcrmodsWrapper, cell, "TonTID");
                                }
                                break;
                            case "GenreInhalt":
                                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    modsClassification(mcrmodsWrapper, cell, "GenreInhalt");
                                }
                                break;
                            case "AnfangEnde":
                                modsNote(mcrmodsWrapper, cell, "start_end");
                                break;
                            case "DBNummerNeu":
                                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    modsIdentifier(mcrmodsWrapper, "intern", cell.getStringCellValue());
                                }
                                break;
                            case "Präsentation":
                                modsClassification(mcrmodsWrapper, cell, "Praesentation");
                                break;
                            case "LängeKopie":
                                String date = cell.getDateCellValue().toString();
                                String[] parts = date.split(" ");
                                String copyTime = parts[3];
                                copyTime = copyTime + ".000";
                                modsPhysicalDescription(mcrmodsWrapper, copyTime);
                                break;
                            case "Betriebsart":
                                modsClassification(mcrmodsWrapper, cell, "Betriebsarten");
                                break;
                            case "Urtitel":
                                modsTitle(mcrmodsWrapper, cell, null, null);
                                break;
                            case "Archivnummer":
                                modsIdentifier(mcrmodsWrapper, "archives", cell.getStringCellValue());
                                break;
                            case "Produktionsnummer":
                                modsIdentifier(mcrmodsWrapper, "production", cell.getStringCellValue());
                                break;
                            case "Redaktion":
                                modsNameEditor(mcrmodsWrapper, cell, "corporate");
                                break;
                            case "Sendereihe":
                                modsNote(mcrmodsWrapper, cell, "serial");
                                break;
                            case "Sprache":
                                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    modsClassification(mcrmodsWrapper, cell, "Sprachen");
                                }
                                break;
                            case "OrgTonträger":
                                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    modsClassification(mcrmodsWrapper, cell, "OrgTonTID");
                                }
                                break;
                            case "Schriftverweis":
                                modsRelatedItem(mcrmodsWrapper, cell, "references");
                                break;
                            case "OrgDatenformat":
                                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    modsClassification(mcrmodsWrapper, cell, "Datenformat");
                                }
                                break;
                            case "DatenformatSichtung":
                                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    modsClassification(mcrmodsWrapper, cell, "DFSichtung");
                                }
                                break;
                            case "DatenformatArchiv":
                                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    cell.setCellType(CellType.STRING);
                                    modsClassification(mcrmodsWrapper, cell, "DFArchiv");
                                }
                                break;
                            case "DOI":
                                String identifier = cell.getStringCellValue();
                                String[] identifierParts = identifier.split(".org");
                                String doi = identifierParts[1];
                                doi = doi.substring(1);
                                modsIdentifier(mcrmodsWrapper, "doi", doi);
                                break;
                            case "Userfeld":
                                modsNote(mcrmodsWrapper, cell, "user");
                                break;
                            case "Aufnahmeort":
                                String dateCaptured;
                                String dateBroadcast;
                                String place;
                                DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSS");

                                dateCaptured = getCellContent(tableHeaderMap, row, "Aufnahmedatum")
                                    .map(XSSFCell::getDateCellValue)
                                    .map(Date::toInstant)
                                    .map(s -> s.atZone(ZoneId.systemDefault()))
                                    .map(ZonedDateTime::toLocalDate)
                                    .map(LocalDate::atStartOfDay)
                                    .map(dateTimeFormatter::format)
                                    .orElse("");


                                dateBroadcast = getCellContent(tableHeaderMap, row, "DatumErstsendung")
                                    .map(XSSFCell::getDateCellValue)
                                    .map(Date::toInstant)
                                    .map(s -> s.atZone(ZoneId.systemDefault()))
                                    .map(ZonedDateTime::toLocalDate)
                                    .map(LocalDate::atStartOfDay)
                                    .map(dateTimeFormatter::format)
                                    .orElse("");

                                place = cell.getStringCellValue();
                                modsOriginInfo(mcrmodsWrapper, dateCaptured, dateBroadcast, "creation", place);
                                break;
                            case "SWSach":
                                String swSachContent = cell.getStringCellValue();
                                String[] titlepartsSach = swSachContent.split("; ");
                                modsSubject(mcrmodsWrapper, null, "topic", "SWSach",
                                    checkSWTitle(swSach, titlepartsSach));
                                break;
                            case "SWWerke":
                                String swWerkeContent = cell.getStringCellValue();
                                String[] titlepartsWerke = swWerkeContent.split("; ");
                                modsSubject(mcrmodsWrapper, "de", "titleInfo", "SWWerke",
                                    checkSWTitle(swWerke, titlepartsWerke));
                                break;
                            case "SWGeo":
                                String swGeoContent = cell.getStringCellValue();
                                String[] titlepartsGeo = swGeoContent.split("; ");
                                modsSubject(mcrmodsWrapper, null, "geographic", "GeoDaten",
                                    checkSWTitle(geoDaten, titlepartsGeo));
                                break;
                            case "SWZeit":
                                LOGGER.info("Hallo welt!");
                                String swZeitContent = cell.getStringCellValue();
                                String[] titlepartsZeit = swZeitContent.split("; ");
                                modsSubject(mcrmodsWrapper, null, "temporal", "SWZeit",
                                    checkSWTitle(zeitDaten, titlepartsZeit));
                                break;
                        }
                    } else if ("Mediensignatur".equals(columnName)) {
                        modsGenre(mcrmodsWrapper, "audio", "sound recording");

                    } else if ("Aufnahmeort".equals(columnName)) {
                        String dateCaptured;
                        String dateBroadcast;
                        String place = "";
                        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSS");

                        dateCaptured = getCellContent(tableHeaderMap, row, "Aufnahmedatum")
                            .map(XSSFCell::getDateCellValue)
                            .map(Date::toInstant)
                            .map(s -> s.atZone(ZoneId.systemDefault()))
                            .map(ZonedDateTime::toLocalDate)
                            .map(LocalDate::atStartOfDay)
                            .map(localDateTime->localDateTime.withNano(0))
                            .map(dateTimeFormatter::format)
                            .orElse("");

                        dateBroadcast = getCellContent(tableHeaderMap, row, "DatumErstsendung")
                            .map(XSSFCell::getDateCellValue)
                            .map(Date::toInstant)
                            .map(s -> s.atZone(ZoneId.systemDefault()))
                            .map(ZonedDateTime::toLocalDate)
                            .map(LocalDate::atStartOfDay)
                            .map(localDateTime->localDateTime.withNano(0))
                            .map(dateTimeFormatter::format)
                            .orElse("");

                        modsOriginInfo(mcrmodsWrapper, dateCaptured, dateBroadcast, "creation", place);
                    }

                }
                String time = getCellContent(tableHeaderMap, row, "OrgLänge")
                    .map(XSSFCell::getDateCellValue)
                    .map(t -> {
                        String[] dateParts = t.toString().split(" ");
                        String originTime = dateParts[3];
                        return originTime + ".000";
                    }).orElse(null);


                String qualityInfo = getCellContent(tableHeaderMap, row, "Tonqualität")
                    .map(XSSFCell::getStringCellValue)
                    .orElseGet(() -> "");

                Boolean orgAnalog = getCellContent(tableHeaderMap, row, "OrgAnalog")
                    .map(XSSFCell::getBooleanCellValue)
                    .orElse(false);

                modsPhysicalDescription(mcrmodsWrapper, time, qualityInfo, orgAnalog);

                modsClassification(mcrmodsWrapper, getGGIDContent(institutionen, i), "Sender");
                String ggID = getCellContent(tableHeaderMap, row, "GGID")
                    .map(XSSFCell::getStringCellValue)
                    .orElse("");
                modsName(mcrmodsWrapper, proccessSWPerson(swPersonen,personenNeu, ggID), "personal");
                try {
                    modsName(mcrmodsWrapper, processRoleTable(personenAlle, personenNeu, personenRollen), "personal", ggID);
                } catch (Exception ex) {
                    LOGGER.warn(ex);
                }

                modsAccessCondition(mcrmodsWrapper, "use and reproduction");
                modsAccessCondition(mcrmodsWrapper, "restriction on access");
                //modsIdentifier(mcrmodsWrapper,"citekey", id);
                //System.out.println(mcrmodsWrapper.getMCRObject().createXML());
                saveFile(mcrmodsWrapper, savePath);
              /*  XMLOutputter xout = new XMLOutputter(Format.getPrettyFormat());
                xout.output(mcrmodsWrapper.getMCRObject().createXML(), System.out);
                activate for testing */
            }
        }
    }

    /**
     * This method return any cell value from the given table
     * @param tableHeaderMap Map with the names of the column and there values
     * @param row   a single row of the excel table, starting with 0
     * @param columnName  the name of the column in witch to be searched
     * @return the cell value if not null
     */
    private Optional<XSSFCell> getCellContent(Map<Integer, String> tableHeaderMap, XSSFRow row, String columnName) {
        return tableHeaderMap.entrySet().stream()
            .filter(entry -> entry.getValue().equals(columnName))
            .findFirst()
            .map(Map.Entry::getKey)
            .map(row::getCell)
            .filter(Objects::nonNull);
    }

    private  Cell getGGIDContent(XSSFSheet table, Integer id) {
        Cell value = null;
        Map<Integer,String> tableHeaderMap = new HashMap<>();
        XSSFRow rowHeader = table.getRow(0);
        for (Cell c : rowHeader) {
            tableHeaderMap.put(c.getColumnIndex(),c.getStringCellValue());
        }
            XSSFRow row = table.getRow(id);
            for (int columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                String columnName = tableHeaderMap.get(columnIndex);
                if (cell != null) {
                    if (columnName.contains("OrgID")) {
                        value = row.getCell(columnIndex +1);
                    }
                }
            }
        return value;
    }

    private Map<String, List<RolePersonTuple>> processRoleTable(XSSFSheet personAll, XSSFSheet personNeu, XSSFSheet personenRollen){
        Map<String,Integer> tableHeaderMap = new HashMap<>();
        XSSFRow rowHeader = personAll.getRow(0);
        for (Cell c : rowHeader) {
            tableHeaderMap.put(c.getStringCellValue(),c.getColumnIndex());
        }

        HashMap<String, List<RolePersonTuple>> result = new HashMap<>();
        Map<String, Person> persons =  processPersonTable(personNeu);
        for (int rowIndex = 1; rowIndex < personAll.getPhysicalNumberOfRows(); rowIndex++) {
            XSSFRow row = personAll.getRow(rowIndex);

            Cell orgDatCell = row.getCell(tableHeaderMap.get("OrgDat"));
            Cell roleCell= row.getCell(tableHeaderMap.get("Rolle"));
            Cell persIDCell= row.getCell(tableHeaderMap.get("PersonenID"));
            if (persIDCell == null){
                LOGGER.warn("Missing PersonID: " + persIDCell);
                continue;
            }
            if (roleCell == null){
                LOGGER.warn("Missing Role: " + roleCell);
                continue;
            }
            orgDatCell.setCellType(CellType.STRING);
            persIDCell.setCellType(CellType.STRING);

            List<RolePersonTuple> personWithRoles = result
                .computeIfAbsent(orgDatCell.getStringCellValue(), (x) -> new ArrayList<>());
            Person person = persons.get(persIDCell.getStringCellValue());
            String role = getPersonRole(personenRollen, roleCell.getStringCellValue());

            personWithRoles.add(new RolePersonTuple(role,person));

        }

        return result;
    }

    private Map<String, String> proccessSWPerson(XSSFSheet swPersons, XSSFSheet personNeu, String id) {
        Map<String, String> swPerson = new HashMap<>();
        Map<String,Integer> tableHeaderMap = new HashMap<>();
        XSSFRow rowHeader = swPersons.getRow(0);
        for (Cell c : rowHeader) {
            tableHeaderMap.put(c.getStringCellValue(),c.getColumnIndex());
        }
        Map<String, Person> persons =  processPersonTable(personNeu);
        for (int rowIndex = 1; rowIndex < swPersons.getPhysicalNumberOfRows(); rowIndex++) {
            XSSFRow row = swPersons.getRow(rowIndex);
            Cell swGGIDCell = row.getCell(tableHeaderMap.get("GGID"));
            swGGIDCell.setCellType(CellType.STRING);
            if(swGGIDCell.getStringCellValue().equals(id)) {
                Cell swPersonIDCell = row.getCell(tableHeaderMap.get("SWPersonID"));
                swPersonIDCell.setCellType(CellType.STRING);
                String stringCellValue = swPersonIDCell.getStringCellValue();
                if (persons.containsKey(stringCellValue)) {
                    Person person = persons.get(stringCellValue);
                    swPerson.put(person.getsName(), person.getfName());
                }
            }
        }

        return swPerson;
    }

    // String == PersonenID
    private Map<String, Person> processPersonTable(XSSFSheet persons){
        Map<String, Person> personContent = new HashMap<>();
        Map<String,Integer> tableHeaderMap = new HashMap<>();
        XSSFRow rowHeader = persons.getRow(0);
        for (Cell c : rowHeader) {
            tableHeaderMap.put(c.getStringCellValue(),c.getColumnIndex());
        }
        for (int rowIndex = 1; rowIndex < persons.getPhysicalNumberOfRows(); rowIndex++) {
            XSSFRow row = persons.getRow(rowIndex);

            row.getCell(tableHeaderMap.get("PersonenID")).setCellType(CellType.STRING); // cast into String

            personContent.put(row.getCell(tableHeaderMap.get("PersonenID")).getStringCellValue(),
                getPerson(row, tableHeaderMap));
        }
        return personContent;
    }

    private String getPersonRole(XSSFSheet personRole, String role) {
        Map<String, String> personRoleMap = new HashMap<>();
        String rolePerson = "";
        for (int rowIndex = 1; rowIndex < personRole.getPhysicalNumberOfRows(); rowIndex++) {
            XSSFRow row = personRole.getRow(rowIndex);
            personRoleMap.put(row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue());
        }
        for (Map.Entry<String, String> entry : personRoleMap.entrySet()) {
            if (entry.getKey().contains(role)) {
                rolePerson =  entry.getValue();
            }
        }
        return rolePerson;
    }

    private Person getPerson(XSSFRow personNew,  Map<String,Integer>tableHeaderMap) {
        Cell personNameCell = personNew.getCell(tableHeaderMap.get("Person"));
        String sName = "";
        String fName = "";
        //Cell persDNBIDCell = personRow.getCell(tableHeaderMap.get("DNBID"));
        if(personNameCell != null) {
            String personNameString = personNameCell.getStringCellValue();
            if (personNameString.contains(", ")) {
                String[] nameParts = personNameString.split(", ");
                sName = nameParts[0];
                fName = nameParts[1];
            } else {
                sName = personNameString;
            }
        }
        return new Person(sName, fName, "");//persDNBIDCell.getStringCellValue()
    }

    // matches the name from GrassArchivDaten.xlsx and Bearbeiter.xslx
	// return the matching name as string
	private static String checkEditorName(XSSFSheet editor, String editorID) {
		String editorName = "";
		for (int i = 1; i < editor.getPhysicalNumberOfRows(); i++) {
			XSSFRow row = editor.getRow(i);
			Cell compare = row.getCell(0);
			if (editorID.equals(compare.getStringCellValue())) {
				for (int n = 1; n < row.getLastCellNum(); n++) {
					Cell cell = row.getCell(n);
					editorName = editorName + cell.getStringCellValue() + ",";
				}
			}
		}
		return editorName;
	}

	// matches the title from GrassArchivDaten.xlsx in column SWSach and
	// SWSach.xslx
	// return the matching ID as string
	private static HashMap<String, String> checkSWTitle(XSSFSheet sheet, String[] swTitle) {
		HashSet<String> hashSet = new HashSet<String>(Arrays.asList(swTitle));
		HashMap<String, String> sachSWID = new HashMap<>();
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow row = sheet.getRow(i);
			Cell compare = row.getCell(1);
			if (compare != null) {
					if (hashSet.contains(compare.getStringCellValue())) {
							Cell cellID = row.getCell(0);
							Cell cellTitle = row.getCell(1);
							cellID.setCellType(CellType.STRING);
							String cellIDString = cellID.getStringCellValue();
							String cellTitleString = cellTitle.getStringCellValue();
							sachSWID.put(cellIDString, cellTitleString);
					}
				
			}
		}
		
		
		return sachSWID;
	}

	public static void saveFile(MCRMODSWrapper mw, String path) throws IOException {
        XMLOutputter xout = new XMLOutputter(Format.getPrettyFormat());
        Document xml = mw.getMCRObject().createXML();

        Element service = xml.getRootElement().getChild("service");
        service.removeContent();
        service.addContent(new Element("servstates").setAttribute("class", "MCRMetaClassification")
            .addContent(new Element("servstate").setAttribute("inherited", "0")
            .setAttribute("classid", "state").setAttribute("categid", "published")));

        try (FileOutputStream fos = new FileOutputStream(path)) {
			xout.output(xml,fos);
		}
	}

	public static void modsBuildHeader(MCRMODSWrapper mw, Integer i) {
		Element rootElement = new Element("mods", MCRConstants.MODS_NAMESPACE)
			.setAttribute("version", "3.6");
		mw.setID("ggrass", i);
		mw.setMODS(rootElement);

	}

	/*public static void modsLocation(MCRMODSWrapper mw, String id) {
		Element location = new Element("location", MCRConstants.MODS_NAMESPACE);
		location.addContent(new Element("url", MCRConstants.MODS_NAMESPACE).setAttribute("access", "object in context")
			.setText("http://webdatenbank.grass-medienarchiv.de/receive/" + id));

		mw.addElement(location);
	}*/

	  public static void modsGenre(MCRMODSWrapper mw, String mediaFormat,
	  String typeOfResource){ HashMap<String, String> attributes = new
	  HashMap<>();

	  attributes.put("type", "intern"); attributes.put("authorityURI",
	  "http://webdatenbank.grass-medienarchiv.de/classifications/mir_genres");
	  attributes.put("valueURI",
	  "http://webdatenbank.grass-medienarchiv.de/classifications/mir_genres#" +
	  mediaFormat );

	  mw.setElement("genre", null, attributes);

	  mw.setElement("typeOfResource", typeOfResource);
	  }

	public static void modsTitle(MCRMODSWrapper mw, Cell title, Cell subTitle, String type) {
	  	Element titleInfo = new Element("titleInfo", MCRConstants.MODS_NAMESPACE)
	  						.setAttribute("lang", "de", Namespace.XML_NAMESPACE);
	  	if(type != null) {
			titleInfo.setAttribute("type", type);
		}
	  	titleInfo.addContent(new Element("title", MCRConstants.MODS_NAMESPACE).setText(title.getStringCellValue()));
	  	if(subTitle != null) {
	  		titleInfo.addContent(new Element("subTitle", MCRConstants.MODS_NAMESPACE).setText(subTitle.toString()));
		}
		mw.addElement(titleInfo);
	}

	// this modsTitle_method will be called if the Object has a modsRelatedItem
	//@return titleInfo as jdom2 Element
	public static Element modsTitle(Cell title, Cell subTitle, String type) {
		Element titleInfo = new Element("titleInfo", MCRConstants.MODS_NAMESPACE)
							.setAttribute("lang", "de", Namespace.XML_NAMESPACE);
		if(type != null && !type.equals("reference")) {
			titleInfo.setAttribute("type", type);
		}
		titleInfo.addContent(new Element("title", MCRConstants.MODS_NAMESPACE).setText(title.getStringCellValue()));
		if(subTitle != null) {
			titleInfo.addContent(new Element("subTitle", MCRConstants.MODS_NAMESPACE).setText(subTitle.toString()));
		}
		return titleInfo;
	}

	public static void modsNote(MCRMODSWrapper mw, Cell context, String att) {
	  	Element node = new Element("note", MCRConstants.MODS_NAMESPACE);
	  	node.setAttribute("type", att);
	  	node.setText(context.getStringCellValue());
	  	mw.addElement(node);
	}

	public static void modsAbstract(MCRMODSWrapper mw, Cell text) {
	  	Element modsAbstract = new Element("abstract", MCRConstants.MODS_NAMESPACE);
	  	modsAbstract.setAttribute("lang", "de", Namespace.XML_NAMESPACE);
	  	modsAbstract.setText(text.getStringCellValue());
	  	mw.addElement(modsAbstract);
	}

	public static void modsClassification(MCRMODSWrapper mw, Cell id, String label) {
        if (id.getCellTypeEnum() == CellType.NUMERIC) { id.setCellType(CellType.STRING);}
        // make sure that the numeric number will be displayed without .0
		HashMap<String, String> attributes = new HashMap<>();
        if(label.contains("Sender")){
            attributes.put("authorityURI", "http://webdatenbank.grass-medienarchiv.de/classifications/Institutionen");
            attributes.put("valueURI", "http://webdatenbank.grass-medienarchiv.de/classifications/" + "Institutionen" + "#" + id);
        }
		else {
            attributes.put("authorityURI", "http://webdatenbank.grass-medienarchiv.de/classifications/" + label);
            attributes.put("valueURI", "http://webdatenbank.grass-medienarchiv.de/classifications/" + label + "#" + id);
        }
        attributes.put("displayLabel", label);

		mw.setElement("classification", null, attributes);
	}

	public static void modsIdentifier(MCRMODSWrapper mw, String type, String text) {
	  	Element identifier = new Element("identifier", MCRConstants.MODS_NAMESPACE)
	  						.setAttribute("type", type).setText(text);
	  	mw.addElement(identifier);
	}

	public static void modsPhysicalDescription(MCRMODSWrapper mw, String time) {
		Element description = new Element("physicalDescription", MCRConstants.MODS_NAMESPACE);
		description.addContent(new Element("extent", MCRConstants.MODS_NAMESPACE).setAttribute("unit", "length")
			.setText(time));
		description.addContent(new Element("reformattingQuality", MCRConstants.MODS_NAMESPACE)
			.setText("preservation"));
		mw.addElement(description);
	}

	public static void modsPhysicalDescription(MCRMODSWrapper mw, String time, String text, boolean orgAnalog) {
		Element description = new Element("physicalDescription", MCRConstants.MODS_NAMESPACE);
		if (time != null) {
			description.addContent(new Element("extent", MCRConstants.MODS_NAMESPACE)
				.setAttribute("unit", "lenght").setText(time));
		}
		if(text != null && !text.isEmpty()) {
			description.addContent(new Element("note", MCRConstants.MODS_NAMESPACE)
				.setAttribute("type", "quality").setText(text));
		}
		if (orgAnalog) {
			description.addContent(new Element("digitalOrigin", MCRConstants.MODS_NAMESPACE)
			.setText("reformatted digital"));
		} else {
            description.addContent(new Element("digitalOrigin", MCRConstants.MODS_NAMESPACE)
                .setText("born digital"));
        }

		if(description.getContentSize()>0){
            mw.addElement(description);
        }
	}

	public static void modsName(MCRMODSWrapper mw, Map<String, List<RolePersonTuple>> processRoleTable, String type, String id) {
        List<RolePersonTuple> rolePersonTuplesList = processRoleTable.get(id);
        Map<String, List<String>> roleMappingTable = new HashMap<>();

        List<String> roleMappingContributor = new ArrayList<>(Arrays.asList("Contributor", "ctb"));
        List<String> roleMappingAutor = new ArrayList<>(Arrays.asList("Author", "aut"));
        List<String> roleMappingInterviewee = new ArrayList<>(Arrays.asList("Interviewee", "ive"));
        List<String> roleMappingSpeaker = new ArrayList<>(Arrays.asList("Speaker", "spk"));
        List<String> roleMappingDirector = new ArrayList<>(Arrays.asList("Director", "drt"));
        List<String> roleMappingRedactor = new ArrayList<>(Arrays.asList("Redactor", "red"));
        List<String> roleMappingTranslator = new ArrayList<>(Arrays.asList("Translator", "trl"));
        List<String> roleMappingOther = new ArrayList<>(Arrays.asList("Other", "oth"));

        roleMappingTable.put("MitwirkendeR", roleMappingContributor);
        roleMappingTable.put("AutorIn", roleMappingAutor);
        roleMappingTable.put("InterviewerIn", roleMappingInterviewee);
        roleMappingTable.put("SprecherIn", roleMappingSpeaker);
        roleMappingTable.put("RegisseurIn", roleMappingDirector);
        roleMappingTable.put("RedakteurIn", roleMappingRedactor);
        roleMappingTable.put("ÜbersetzerIn", roleMappingTranslator);
        roleMappingTable.put("---", roleMappingOther);

        for (RolePersonTuple personTuple : rolePersonTuplesList) {
            Element name = new Element("name", MCRConstants.MODS_NAMESPACE).setAttribute("type", type);
            String family = personTuple.person.getsName();
            String given = personTuple.person.getfName();
            List<String> role = roleMappingTable.get(personTuple.role);
            if (!family.isEmpty()) {
                name.addContent(new Element("namePart", MCRConstants.MODS_NAMESPACE)
                    .setAttribute("type", "family").setText(family));
            }
            name.addContent(new Element("namePart", MCRConstants.MODS_NAMESPACE)
                .setAttribute("type", "given").setText(given));
            Element modsRole = new Element("role", MCRConstants.MODS_NAMESPACE)
                .addContent(new Element("roleTerm", MCRConstants.MODS_NAMESPACE)
                    .setAttribute("type", "code")
                    .setText(role.get(1)));
     // AKN     .addContent(new Element("roleTerm", MCRConstants.MODS_NAMESPACE)
     // AKN         .setAttribute("type", "text")
     // AKN         .setText(role.get(0)));
            name.addContent(modsRole);
            mw.addElement(name);
        }
	}

	public static void modsName(MCRMODSWrapper mw, Map<String, String>swPerson, String type) {
	      for (Map.Entry<String, String> entry : swPerson.entrySet()) {
              Element modsSubject = new Element("subject", MCRConstants.MODS_NAMESPACE);
              Element name = new Element("name", MCRConstants.MODS_NAMESPACE).setAttribute("type", type);
              if (!entry.getKey().isEmpty()) {
                  name.addContent(new Element("namePart", MCRConstants.MODS_NAMESPACE)
                      .setAttribute("type", "family").setText(entry.getKey()));
              }
              name.addContent(new Element("namePart", MCRConstants.MODS_NAMESPACE)
                  .setAttribute("type", "given").setText(entry.getValue()));
              modsSubject.addContent(name);
              mw.addElement(modsSubject);
          }
    }


	public static void modsNameEditor(MCRMODSWrapper mw, Cell editor, String type) {
	  	Element modsName = new Element("name", MCRConstants.MODS_NAMESPACE);
	  	if (type.contains("corporate")){ // in this case name check if the type contains corporate
	  		modsName.setAttribute("type", type);
			modsName.addContent(new Element("displayForm", MCRConstants.MODS_NAMESPACE).setText(editor.getStringCellValue()));
			modsName.addContent(modsRole(
				"marcrelator", "code", "red", "Redactor"));
		}
	  	else if(type.contains("unbekannt")) {
	  		String[] parts = type.split(",");
	  		String unknown = parts[0];
	  		String initials = parts[1];
	  		modsName.setAttribute("type", "personal");
	  		modsName.addContent(new Element("displayForm", MCRConstants.MODS_NAMESPACE).setText(unknown + " (" + initials + ")"));
			modsName.addContent(new Element("nameIdentifier", MCRConstants.MODS_NAMESPACE)
				.setAttribute("type", "intern").setText(editor.getStringCellValue()));
			modsName.addContent(modsRole(
				"marcrelator", "code", "mdc", "Metadata contact"));
		} else {
			String[] parts = type.split(",");
			String given = parts[1];
			given = given.trim();
			String family = parts[0];
			modsName.setAttribute("type", "personal");
			modsName.addContent(new Element("namePart", MCRConstants.MODS_NAMESPACE).setAttribute("type", "given").setText(given));
			modsName.addContent(new Element("namePart", MCRConstants.MODS_NAMESPACE).setAttribute("type", "family").setText(family));
			modsName.addContent(new Element("nameIdentifier", MCRConstants.MODS_NAMESPACE)
				.setAttribute("type", "intern").setText(editor.getStringCellValue()));
			modsName.addContent(modsRole(
				"marcrelator", "code", "mdc", "Metadata contact"));
		}

		mw.addElement(modsName);
	}
    // Helper_class for modsName
	// @return role Element
	public static Element modsRole(String authority, String type1, String text1, String text2) {
	  	Element role = new Element("role", MCRConstants.MODS_NAMESPACE);
	  	role.addContent(new Element("roleTerm", MCRConstants.MODS_NAMESPACE)
			.setAttribute("authority", authority).setAttribute("type", type1).setText(text1));
	// AKN 	role.addContent(new Element("roleTerm", MCRConstants.MODS_NAMESPACE)
	// AKN 	.setAttribute("authority", authority).setText(text2));
	  	return role;
	}

	public static void modsRelatedItem(MCRMODSWrapper mw, Cell text, String type) {
		Element relatedItem = new Element("relatedItem", MCRConstants.MODS_NAMESPACE).setAttribute("type", type);
		relatedItem.addContent(modsTitle(text, null, "reference"));
        mw.addElement(relatedItem);
	}


    public static void modsOriginInfo(MCRMODSWrapper mw, String dateCaptured, String dateBroadcast, String eventType,
        String place) {
            Element originInfo = new Element("originInfo", MCRConstants.MODS_NAMESPACE).setAttribute("eventType", eventType);
            if(!dateBroadcast.isEmpty())
            originInfo.addContent(new Element("dateIssued", MCRConstants.MODS_NAMESPACE).setAttribute("encoding", "w3cdtf")
            .setText(dateBroadcast));
            if(!dateCaptured.isEmpty())
            originInfo.addContent(new Element("dateCaptured", MCRConstants.MODS_NAMESPACE).setAttribute("encoding", "w3cdtf")
            .setText(dateCaptured));
            if(!place.isEmpty()) {
                Element modsPlace = new Element("place", MCRConstants.MODS_NAMESPACE);
                modsPlace.addContent(new Element("placeTerm", MCRConstants.MODS_NAMESPACE).setAttribute("type", "text")
                    .setText(place));
                originInfo.addContent(modsPlace);
            }
            if(originInfo.getContentSize()>0){
                mw.addElement(originInfo);
            }
    }

	public static void modsSubject(MCRMODSWrapper mw, String language, String modsName, String type, HashMap<String, String> values){
	      if(language == null) {
	          for (Map.Entry<String, String> entry : values.entrySet()){
	              Element subject = new Element("subject", MCRConstants.MODS_NAMESPACE);
	              subject.addContent(new Element(modsName, MCRConstants.MODS_NAMESPACE).setAttribute("authorityURI",
                      "http://webdatenbank.grass-medienarchiv.de/classifications/" + type).setAttribute("valueURI",
                      "http://webdatenbank.grass-medienarchiv.de/classifications/" + type + "#" + entry.getKey())
                      .setText(entry.getValue()));
                  mw.addElement(subject);
              }
          } else {
              values.entrySet().stream().sorted((e1, e2) -> ((Integer) Integer.parseInt(e1.getKey()))
                  .compareTo((Integer) Integer.parseInt(e2.getKey()))).forEach((Map.Entry<String, String> entry) -> {
                      Element subject = new Element("subject", MCRConstants.MODS_NAMESPACE);
                      subject.addContent(new Element(modsName, MCRConstants.MODS_NAMESPACE).setAttribute(
                          "authorityURI", "http://webdatenbank.grass-medienarchiv.de/classifications/" + type)
                          .setAttribute("lang", language, MCRConstants.XML_NAMESPACE).setAttribute("valueURI",
                              "http://webdatenbank.grass-medienarchiv.de/classifications/" + type + "#" + entry.getKey())
                          .addContent(new Element("title", MCRConstants.MODS_NAMESPACE).setText(entry.getValue()))
                      );
                  mw.addElement(subject);
              });
          }
    }

	public static void modsAccessCondition(MCRMODSWrapper mw, String type) {
        if(type.contains("use and reproduction")) {
            Element accessCondition = new Element("accessCondition", MCRConstants.MODS_NAMESPACE)
                .setAttribute("type", type).setAttribute("href",
                    "http://webdatenbank.grass-medienarchiv.de/classifications/mir_rights#rights_reserved",
                    MCRConstants.XLINK_NAMESPACE);
            mw.addElement(accessCondition);
        }
        if(type.contains("restriction on access")) {
            Element accessCondition = new Element("accessCondition", MCRConstants.MODS_NAMESPACE)
                .setAttribute("type", type).setAttribute("href",
                    "http://webdatenbank.grass-medienarchiv.de/classifications/mir_access#intern",
                    MCRConstants.XLINK_NAMESPACE);
            mw.addElement(accessCondition);
        }
    }
}
