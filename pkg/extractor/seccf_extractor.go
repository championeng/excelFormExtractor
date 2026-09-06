package extractor

import (
	"encoding/json"
	"fmt"
	"reflect"
	"strings"

	"github.com/championeng/excelFormExtractor/pkg/utils"
	"github.com/xuri/excelize/v2"
)

// SearchCriteria defines what to look for and where
type SearchCriteria struct {
	SearchTerms           []string                   // Multiple possible terms to search for
	CellRanges            []CellRange                // Multiple cell ranges to search in
	DualColumnCheckBoxClf bool                       // check side by side column
	DualColumnClfCriteria DualClassificationCriteria // Add this to map checkbox text to values
	TriColumnCheckBoxClf  bool                       // check tri-side by side column
	TriColumnClfCriteria  TriClassificationCritera   // Add this to map checkbox text to values
	BoolCheckBox          bool
	BoolClfCriteria       BoolClassificationCriteria
	BoolContainsImage     bool
	BoolClfContainsImage  BoolClassificationCriteria
	Offset                int // Default offset of value for simple fields
}

type ColumnMapping struct {
	FieldName   string
	SearchTerms []string
	FoundColumn string // Will store the actual column letter once found
}

type ClassificationCriteria struct {
	Label       string
	SearchTerms []string // search terms for extra check if form control has that name or not
	Offset      int
}

type BoolClassificationCriteria struct {
	Offset      int
	SearchTerms []string // search terms for extra check if form control has that name or not
}

type DualClassificationCriteria struct {
	TYPE_1 ClassificationCriteria
	TYPE_2 ClassificationCriteria
}

type TriClassificationCritera struct {
	TYPE_1 ClassificationCriteria
	TYPE_2 ClassificationCriteria
	TYPE_3 ClassificationCriteria
}

// CellRange represents an Excel cell range
type CellRange struct {
	StartCell string // e.g., "B12"
	EndCell   string // e.g., "D12"
}

type BuyerDetails struct {
	SheetName                       string `json:"sheet_name"`
	PartNumber                      string `json:"part_number"`
	PartDescription                 string `json:"part_description"`
	ClassificationOfItem            string `json:"classification_of_item"`
	ControlListClassificationNumber string `json:"control_list_classification_number"`
	RFQ                             string `json:"rfq"`
	BuildToPrint                    bool   `json:"build_to_print"`
	ManufacturedToSpecification     bool   `json:"manufactured_to_specification"`
	OriginalEquipmentManufacturer   bool   `json:"original_equipment_manufacturer"`
	Modified                        bool   `json:"modified"`
}

type ProductDetails struct {
	// Sheet Metadata
	SheetName string `json:"sheet_name"`

	// Supplier Details
	SupplierPartNumber    string `json:"supplier_part_number"`
	SupplierCompanyName   string `json:"supplier_company_name"`
	SupplierFullAddress   string `json:"supplier_full_address"`
	SupplierCountry       string `json:"supplier_country"`
	SupplierCompanyNumber string `json:"supplier_company_number"`

	// Manufacturer Details
	ManufacturerPartNumber    string `json:"manufacturer_part_number"`
	ManufacturerCompanyName   string `json:"manufacturer_company_name"`
	ManufacturerFullAddress   string `json:"manufacturer_full_address"`
	ManufacturerCountry       string `json:"manufacturer_country"`
	ManufacturerCompanyNumber string `json:"manufacturer_company_number"`

	// Product Details
	CountryOfOrigin                 string `json:"country_of_origin"`
	CustomsTariffCode               string `json:"customs_tariff_code"`
	ExportControlRegulated          string `json:"export_control_regulated"` // Yes/No
	PartClassification              string `json:"part_classification"`      // DU, MIL, CIVIL
	ControlListClassificationNumber string `json:"control_list_classification_number"`
	ThirdCountryControlledContent   string `json:"third_country_controlled_content"` // Yes/No
	EndUserStatementRequired        string `json:"end_user_statement_required"`      // Yes/No
	ExportLicenceShipmentRequired   string `json:"export_licence_shipment_required"` // Yes/No
	ExportLicenceEndUserRequired    string `json:"export_licence_end_user_required"` // Yes/No/End user not advised to supplier
	AdditionalExportDocsRequired    string `json:"additional_export_docs_required"`  // Yes/No
	// AdditionalShipmentRequirements  string `json:"additional_shipment_requirements"`

	// Mandatory
	TransferReexportConditions string `json:"transfer_reexport_conditions"`

	// Supplier Representative
	RepresentativeName      string `json:"representative_name"`
	RepresentativePosition  string `json:"representative_position"`
	RepresentativeSignature bool   `json:"representative_signature"` // Available/Not Available
	SupplierCompanySeal     string `json:"supplier_company_seal"`    // Available/Not Available
	SignatureDate           string `json:"signature_date"`           // Date format
}

type ControlCotent struct {
	SheetName                       string `json:"sheet_name"`
	ItemNum                         string `json:"item_num"`
	PartNumber                      string `json:"part_number"`
	ComponentManufacturerPartNumber string `json:"component_manufacturer_part_number"`
	PartDescription                 string `json:"part_description"`
	ManufacturerOfComponent         string `json:"manufacturer_of_component"`
	ExportRegulationCountry         string `json:"export_regulation_country"`
	DualControlListClfNum           string `json:"dual_control_list_clf_num"`
	MilitaryControlListClfNum       string `json:"military_control_list_clf_num"`
	IndicateLicenseApplication      string `json:"inidcate_license_application"`
	TopLevelDeliverableItem         string `json:"top_level_delierable_item"`
	USML_N                          string `json:"usml_n"`
	ECCN_N                          string `json:"eccn_n"`
	US_EA_CONTENT_RATIO             string `json:"us_ea_content_ratio"`
}

// Generic interface for structures with SheetName
type SheetNameGetter interface {
	GetSheetName() string
}

// Add methods to both structs to implement SheetNameGetter
func (b *BuyerDetails) GetSheetName() string {
	return b.SheetName
}

func (p *ProductDetails) GetSheetName() string {
	return p.SheetName
}

type SECCFExtraction struct {
	BuyerDetails      *BuyerDetails   `json:"buyer_details"`
	ProductDetails    *ProductDetails `json:"product_details"`
	ControlledContent []ControlCotent `json:"controlled_content"`
	// add more extraction if possible
}

type ExcelExtractor struct {
	file         *excelize.File
	companyNames []string
	Extraction   *SECCFExtraction
}

// //////////////////////////
// // Specific extractor ////
// //////////////////////////

// Add a method to handle value extraction based on criteria type
type ValueExtractor interface {
	Extract(e *ExcelExtractor, sheetName string, criteria SearchCriteria, cellRange CellRange) (interface{}, error)
}

// Implement different extractors for different types of fields
type SimpleValueExtractor struct{}
type BoolCheckBoxExtractor struct{}
type DualColumnClfExtractor struct{}
type TriColumnClfExtractor struct{}
type BoolContainsImageExtractor struct{}

func (s *SimpleValueExtractor) Extract(e *ExcelExtractor, sheetName string, criteria SearchCriteria, cellRange CellRange) (interface{}, error) {
	adjacentRange := getAdjacentRange(cellRange, criteria.Offset)
	return e.GetCellValue(adjacentRange, sheetName)
}

func (c *BoolCheckBoxExtractor) Extract(e *ExcelExtractor, sheetName string, criteria SearchCriteria, cellRange CellRange) (interface{}, error) {
	cell := getAdjacentRange(cellRange, criteria.BoolClfCriteria.Offset).StartCell
	return e.isCheckBoxChecked(sheetName, cell, criteria.BoolClfCriteria.SearchTerms)
}

func (c *BoolContainsImageExtractor) Extract(e *ExcelExtractor, sheetName string, criteria SearchCriteria, cellRange CellRange) (interface{}, error) {
	cell := getAdjacentRange(cellRange, criteria.BoolClfContainsImage.Offset).StartCell
	return e.doesContainImage(sheetName, cell)
}

func (d *DualColumnClfExtractor) Extract(e *ExcelExtractor, sheetName string, criteria SearchCriteria, cellRange CellRange) (interface{}, error) {
	cellType1 := getAdjacentRange(cellRange, criteria.DualColumnClfCriteria.TYPE_1.Offset).StartCell
	cellType2 := getAdjacentRange(cellRange, criteria.DualColumnClfCriteria.TYPE_2.Offset).StartCell
	// fmt.Println("cellType1", cellType1, " cellType2", cellType2)

	isType1, err := e.isCheckBoxChecked(sheetName, cellType1, criteria.DualColumnClfCriteria.TYPE_1.SearchTerms)
	if err != nil {
		fmt.Printf("Error checking %s classification: %v\n", criteria.DualColumnClfCriteria.TYPE_1.Label, err)
		return "", err
	}

	isType2, err := e.isCheckBoxChecked(sheetName, cellType2, criteria.DualColumnClfCriteria.TYPE_2.SearchTerms)
	if err != nil {
		fmt.Printf("Error checking %s classification: %v\n", criteria.DualColumnClfCriteria.TYPE_2.Label, err)
		return "", err
	}

	if isType1 && !isType2 {
		return criteria.DualColumnClfCriteria.TYPE_1.Label, nil
	} else if !isType1 && isType2 {
		return criteria.DualColumnClfCriteria.TYPE_2.Label, nil
	}
	return "", nil
}

func (d *TriColumnClfExtractor) Extract(e *ExcelExtractor, sheetName string, criteria SearchCriteria, cellRange CellRange) (interface{}, error) {
	cellType1 := getAdjacentRange(cellRange, criteria.TriColumnClfCriteria.TYPE_1.Offset).StartCell
	cellType2 := getAdjacentRange(cellRange, criteria.TriColumnClfCriteria.TYPE_2.Offset).StartCell
	cellType3 := getAdjacentRange(cellRange, criteria.TriColumnClfCriteria.TYPE_3.Offset).StartCell

	isType1, err := e.isCheckBoxChecked(sheetName, cellType1, criteria.TriColumnClfCriteria.TYPE_1.SearchTerms)
	if err != nil {
		fmt.Printf("Error checking %s classification: %v\n", criteria.TriColumnClfCriteria.TYPE_1.Label, err)
		return "", err
	}

	isType2, err := e.isCheckBoxChecked(sheetName, cellType2, criteria.TriColumnClfCriteria.TYPE_2.SearchTerms)
	if err != nil {
		fmt.Printf("Error checking %s classification: %v\n", criteria.TriColumnClfCriteria.TYPE_2.Label, err)
		return "", err
	}

	isType3, err := e.isCheckBoxChecked(sheetName, cellType3, criteria.TriColumnClfCriteria.TYPE_3.SearchTerms)
	if err != nil {
		fmt.Printf("Error checking %s classification: %v\n", criteria.TriColumnClfCriteria.TYPE_3.Label, err)
		return "", err
	}

	if isType1 && !isType2 && !isType3 {
		return criteria.TriColumnClfCriteria.TYPE_1.Label, nil
	} else if !isType1 && isType2 && !isType3 {
		return criteria.TriColumnClfCriteria.TYPE_2.Label, nil
	} else if !isType1 && !isType2 && isType3 {
		return criteria.TriColumnClfCriteria.TYPE_3.Label, nil
	}
	return "", nil
}

// ///////////////////////////////
// // Specific extractor ENDS ////
// ///////////////////////////////

// Helper function to set values using reflection
func setValue(field reflect.Value, value interface{}) {
	switch field.Kind() {
	case reflect.String:
		field.SetString(value.(string))
	case reflect.Bool:
		field.SetBool(value.(bool))
		// Add more types as needed
	}
}

func (e *ExcelExtractor) ReplaceCompanyNames(items []string) []string {
	var results []string

	// Replace each company name in the string
	for _, companyName := range e.companyNames {
		for _, item := range items {
			/// Replace the placeholder with the company name
			replaced := strings.ReplaceAll(item, "{companyName}", companyName)
			results = append(results, replaced)
		}
	}
	return results
}

func MakeSECCFExtractor(filePath string, companyNames CompanyNameList) (*ExcelExtractor, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open Excel file: %w", err)
	}

	return &ExcelExtractor{
		file:         f,
		companyNames: companyNames,
		Extraction:   &SECCFExtraction{},
	}, nil
}

func (e *ExcelExtractor) searchSheetName(searchWord string) (bool, string, error) {
	sheetList := e.file.GetSheetList()

	wordFound := false
	foundSheetName := ""

	for _, sheetName := range sheetList {
		// Convert both strings to lowercase for case-insensitive comparison
		if strings.Contains(strings.ToLower(sheetName), strings.ToLower(searchWord)) {
			// fmt.Printf("Found '%s' in sheet name: '%s' at position %d\n", searchWord, sheetName, index+1)
			wordFound = true
			foundSheetName = sheetName
		}
	}

	if !wordFound {
		return wordFound, "", SheetNotFoundError{searchWord: searchWord}
	}
	return wordFound, foundSheetName, nil
}

func (e *ExcelExtractor) GetCellValue(cellRange CellRange, sheetName string) (string, error) {
	// First try getting the value from the start cell
	value, err := e.file.GetCellValue(sheetName, cellRange.StartCell)
	if err != nil {
		return "", fmt.Errorf("failed to get cell value: %w", err)
	}

	// If we got a value, return it
	if value != "" {
		return strings.TrimSpace(value), nil
	}

	// If no value in start cell, check if it's part of a merged range
	mergedCells, err := e.file.GetMergeCells(sheetName)
	if err != nil {
		return "", fmt.Errorf("failed to get merged cells: %w", err)
	}

	// Check each merged range
	for _, mergedCell := range mergedCells {
		if e.isCellInRange(cellRange.StartCell, &mergedCell) {
			// fmt.Printf("Cellvalue: %s\n", strings.TrimSpace(mergedCell.GetCellValue()))
			return strings.TrimSpace(mergedCell.GetCellValue()), nil
		}
	}

	return "", nil
}

// isCellInRange checks if a cell is within a merged cell range
func (e *ExcelExtractor) isCellInRange(cell string, mergedCell *excelize.MergeCell) bool {
	return strings.HasPrefix(mergedCell.GetStartAxis(), cell) ||
		strings.HasPrefix(cell, mergedCell.GetStartAxis())
}

// getAdjacentRange returns the adjacent cell range
func getAdjacentRange(cellRange CellRange, offset int) CellRange {
	// Parse the column and row from StartCell
	startCol := strings.Split(cellRange.StartCell, "")[0]
	startRow := strings.Split(cellRange.StartCell, "")[1:]

	// Parse the column and row from EndCell
	endCol := strings.Split(cellRange.EndCell, "")[0]
	endRow := strings.Split(cellRange.EndCell, "")[1:]

	// Calculate new columns (offset to the right)
	newStartCol := string(rune(startCol[0] + byte(offset)))
	newEndCol := string(rune(endCol[0] + byte(offset)))

	return CellRange{
		StartCell: newStartCol + strings.Join(startRow, ""),
		EndCell:   newEndCol + strings.Join(endRow, ""),
	}
}

func (e *ExcelExtractor) ToJson() string {
	jsonBytes, err := json.Marshal(e.Extraction)
	if err != nil {
		fmt.Println("Error:", err)
		return string("{}")
	}
	return string(jsonBytes)
}

func (e *ExcelExtractor) findHeaderRow(sheetName string, columnMappings []ColumnMapping) (int, error) {
	// Look for headers between rows 10-12
	for row := 10; row <= 12; row++ {
		// Check if this row contains known headers
		for _, mapping := range columnMappings {
			for _, term := range mapping.SearchTerms {
				value, _ := e.file.GetCellValue(sheetName, fmt.Sprintf("A%d", row))
				if strings.Contains(strings.ToLower(value), strings.ToLower(term)) {
					return row, nil
				}
			}
		}
	}
	return 11, fmt.Errorf("header row not found")
}

func (e *ExcelExtractor) findColumnByHeader(sheetName string, headerRow int, searchTerms []string) (string, error) {
	// Get all cells in the header row
	cols, err := e.file.GetCols(sheetName)
	if err != nil {
		return "", fmt.Errorf("failed to get columns: %w", err)
	}

	// Look through each column
	for colIdx, col := range cols {
		if len(col) >= headerRow {
			headerCell := utils.RemoveExtraSpaces(strings.TrimSpace(strings.ToLower(col[headerRow-1])))
			// Check if any search term matches
			for _, term := range searchTerms {
				termLower := utils.RemoveExtraSpaces(strings.ToLower(term))
				if strings.Contains(headerCell, termLower) {
					// Convert column index to letter (0 = A, 1 = B, etc.)
					colName, err := excelize.ColumnNumberToName(colIdx + 1)
					fmt.Println("Found colName:", colName, "for term:", term, "headerCell: ", headerCell)
					if err != nil {
						return "", err
					}
					return colName, nil
				}
			}
		}
	}
	return "", fmt.Errorf("column not found for search terms: %v", searchTerms)
}

func (e *ExcelExtractor) extractControlledContent(sheetName string, columnMappings []ColumnMapping) []ControlCotent {
	var contents []ControlCotent

	// Find the header row (assuming it's around row 10-12)
	// headerRow := 11 // You might want to make this dynamic too
	headerRow, _ := e.findHeaderRow(sheetName, columnMappings) // default to 11

	// Find actual columns for each mapping
	for i := range columnMappings {
		col, err := e.findColumnByHeader(sheetName, headerRow, columnMappings[i].SearchTerms)
		if err != nil {
			fmt.Printf("Warning: Could not find column for %s: %v\n", columnMappings[i].FieldName, err)
			continue
		}
		columnMappings[i].FoundColumn = col
	}

	// Start from the row after header
	row := headerRow + 1
	for {
		// Check if row is empty (using first column as indicator)
		if len(columnMappings) == 0 || columnMappings[0].FoundColumn == "" {
			break
		}

		firstCellValue, _ := e.file.GetCellValue(sheetName, fmt.Sprintf("%s%d", columnMappings[0].FoundColumn, row))
		if firstCellValue == "" {
			break
		}

		content := ControlCotent{
			SheetName: sheetName,
		}

		// Use reflection to set fields dynamically
		contentValue := reflect.ValueOf(&content).Elem()

		for _, mapping := range columnMappings {
			if mapping.FoundColumn == "" {
				continue
			}

			cellValue, _ := e.file.GetCellValue(sheetName, fmt.Sprintf("%s%d", mapping.FoundColumn, row))
			field := contentValue.FieldByName(mapping.FieldName)

			if field.IsValid() && field.CanSet() {
				field.SetString(strings.TrimSpace(cellValue))
			}
		}

		contents = append(contents, content)
		row++
	}

	return contents
}

func (e *ExcelExtractor) extractDetails(details interface{}, sheetName string, criteria map[string]SearchCriteria) {
	// Get the reflect.Value of the pointer to the struct
	detailsValue := reflect.ValueOf(details).Elem()

	for fieldName, searchCriteria := range criteria {
		keyCellFound := false
		for _, cellRange := range searchCriteria.CellRanges {
			// Get 'KEY' cell from the potential label cell
			value, err := e.GetCellValue(cellRange, sheetName)
			if err != nil {
				fmt.Printf("Error trying to retrive cell value: %v\n", err)
				return
			}

			// Check if the 'KEY' Cell value matches any of our search terms
			for _, searchTerm := range searchCriteria.SearchTerms {
				if strings.Contains(utils.RemoveExtraSpaces(strings.ToLower(value)), utils.RemoveExtraSpaces(strings.ToLower(searchTerm))) {
					keyCellFound = true
					var extractor ValueExtractor

					// Select appropriate extractor based on criteria type
					if searchCriteria.DualColumnCheckBoxClf {
						extractor = &DualColumnClfExtractor{}
					} else if searchCriteria.TriColumnCheckBoxClf {
						extractor = &TriColumnClfExtractor{}
					} else if searchCriteria.BoolCheckBox {
						extractor = &BoolCheckBoxExtractor{}
					} else if searchCriteria.BoolContainsImage {
						extractor = &BoolContainsImageExtractor{}
					} else {
						extractor = &SimpleValueExtractor{}
					}

					// Extract the value
					extractedValue, err := extractor.Extract(e, sheetName, searchCriteria, cellRange)
					if err != nil {
						fmt.Println("error extracting value: %w", err)
						return
					}

					// Set the field using reflection
					field := detailsValue.FieldByName(fieldName)
					if field.IsValid() && field.CanSet() {
						setValue(field, extractedValue)
					} else {
						fmt.Printf("field: %s is isValid: %v and canSet: %v\n", fieldName, field.IsValid(), field.CanSet())
					}

					break
				}
			}
		}
		if !keyCellFound {
			fmt.Printf("Field %s not found in excel\n", fieldName)
		}
	}
}

func (e *ExcelExtractor) ReadFormControls() {
	sheetName := e.Extraction.ProductDetails.SheetName
	formControls, err := e.file.GetFormControls(sheetName) // sheet name
	if err != nil {
		fmt.Println(err)
		return
	}

	for _, control := range formControls {
		fmt.Printf("[%s] Control Cell %s, Control checked: %v, control.Paragraph %v, control.CurrentVal %v, cellLink: %s, offsetX: %v, offsetY: %v, Control cell text: %s\n", sheetName, control.Cell, control.Checked, control.Paragraph, control.CurrentVal, control.CellLink, control.Format.OffsetX, control.Format.OffsetY, control.Text)

		// if control.Type == excelize.FormControlCheckBox {
		// }
	}
}

func (e *ExcelExtractor) isCheckBoxChecked(sheetName string, cell string, classificationTexts []string) (bool, error) {
	formControls, err := e.file.GetFormControls(sheetName)
	if err != nil {
		return false, fmt.Errorf("failed to get form controls: %w", err)
	}

	for _, control := range formControls {
		if control.Cell == cell {
			if control.Type == excelize.FormControlCheckBox {
				for _, text := range classificationTexts {
					for _, paraText := range control.Paragraph {
						if strings.EqualFold(paraText.Text, text) {
							// fmt.Printf("[%s] Control Cell %s, Control checked: %v control.Paragraph %v, Control cell text: %s, cell %s\n", sheetName, control.Cell, control.Checked, control.Paragraph, control.Text, cell)
							return control.Checked, nil
						}
					}
				}
			}
		}
	}
	return false, nil
}

func (e *ExcelExtractor) doesContainImage(sheetName string, cell string) (bool, error) {
	images, err := e.file.GetPictures(sheetName, cell)
	// fmt.Printf("images: %v cell %s\n", images, cell)
	if err != nil {
		return false, fmt.Errorf("failed to get form controls: %w", err)
	}

	if len(images) == 0 {
		return false, nil
	}
	return true, nil
}

func (e *ExcelExtractor) Extract() SECCFExtraction {
	_, buyerSheetName, err := e.searchSheetName("buyer details")

	buyerDetailsCriteria := map[string]SearchCriteria{
		"PartNumber": {
			SearchTerms: []string{"part number", "part-nr", "part_number"},
			CellRanges: []CellRange{
				{StartCell: "B12", EndCell: "D12"},
			},
			Offset: 3,
		},
		"PartDescription": {
			SearchTerms: []string{"description", "desc", "part description"},
			CellRanges: []CellRange{
				{StartCell: "B13", EndCell: "D13"},
			},
			Offset: 3,
		},
		"ControlListClassificationNumber": {
			SearchTerms: []string{"control list classification number"},
			CellRanges: []CellRange{
				{StartCell: "B18", EndCell: "D18"},
			},
			Offset: 3,
		},
		"RFQ": {
			SearchTerms: []string{"RQF", "quote reference"},
			CellRanges: []CellRange{
				{StartCell: "B19", EndCell: "D19"},
			},
			Offset: 3,
		},
		"BuildToPrint": {
			SearchTerms: []string{"Build To Print"},
			CellRanges: []CellRange{
				{StartCell: "B21", EndCell: "F21"},
			},
			BoolCheckBox: true,
			BoolClfCriteria: BoolClassificationCriteria{
				Offset:      5,
				SearchTerms: []string{"YES"},
			},
		},
		"ManufacturedToSpecification": {
			SearchTerms: []string{"Manufactured to specification", "(MTS)"},
			CellRanges: []CellRange{
				{StartCell: "B22", EndCell: "F22"},
			},
			BoolCheckBox: true,
			BoolClfCriteria: BoolClassificationCriteria{
				Offset:      5,
				SearchTerms: []string{"YES"},
			},
		},
		"OriginalEquipmentManufacturer": {
			SearchTerms: []string{"Original Equipment Manufacturer"},
			CellRanges: []CellRange{
				{StartCell: "B23", EndCell: "F23"},
			},
			BoolCheckBox: true,
			BoolClfCriteria: BoolClassificationCriteria{
				Offset:      5,
				SearchTerms: []string{"YES"},
			},
		},
		"Modified": {
			SearchTerms: []string{"Modified"},
			CellRanges: []CellRange{
				{StartCell: "B25", EndCell: "F25"},
			},
			BoolCheckBox: true,
			BoolClfCriteria: BoolClassificationCriteria{
				Offset:      5,
				SearchTerms: []string{"YES"},
			},
		},
		"ClassificationOfItem": {
			SearchTerms: e.ReplaceCompanyNames([]string{"{companyName} Classification of item"}),
			CellRanges: []CellRange{
				{StartCell: "B15", EndCell: "D15"},
			},
			DualColumnCheckBoxClf: true,
			DualColumnClfCriteria: DualClassificationCriteria{
				TYPE_1: ClassificationCriteria{
					Label:       "DUAL",
					SearchTerms: []string{"Dual", "DU"},
					Offset:      3,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "MILITARY",
					SearchTerms: []string{"Military", "MIL"},
					Offset:      4,
				},
			},
		},
	}

	if err != nil {
		fmt.Println("%w", err)
	} else {
		e.Extraction.BuyerDetails = &BuyerDetails{
			SheetName: buyerSheetName,
		}
		e.extractDetails(e.Extraction.BuyerDetails, buyerSheetName, buyerDetailsCriteria)
	}

	_, productSheetName, err_product_sheet_search := e.searchSheetName("product details")

	productDetailsCriteria := map[string]SearchCriteria{
		"SupplierPartNumber": {
			SearchTerms: []string{"Supplier part number"},
			CellRanges: []CellRange{
				{StartCell: "C11", EndCell: "D11"},
			},
			Offset: 2,
		},
		"SupplierCompanyName": {
			SearchTerms: []string{"company name"},
			CellRanges: []CellRange{
				{StartCell: "C12", EndCell: "C12"},
			},
			Offset: 1,
		},
		"SupplierFullAddress": {
			SearchTerms: []string{"full address"},
			CellRanges: []CellRange{
				{StartCell: "C13", EndCell: "C13"},
			},
			Offset: 1,
		},
		"SupplierCountry": {
			SearchTerms: []string{"Country"},
			CellRanges: []CellRange{
				{StartCell: "C14", EndCell: "C14"},
			},
			Offset: 1,
		},
		"SupplierCompanyNumber": {
			SearchTerms: []string{"company number"},
			CellRanges: []CellRange{
				{StartCell: "C15", EndCell: "C15"},
			},
			Offset: 1,
		},
		"ManufacturerPartNumber": {
			SearchTerms: []string{"manufacturer part number"},
			CellRanges: []CellRange{
				{StartCell: "C16", EndCell: "C116"},
			},
			Offset: 2,
		},
		"ManufacturerCompanyName": {
			SearchTerms: []string{"company name"},
			CellRanges: []CellRange{
				{StartCell: "C17", EndCell: "C17"},
			},
			Offset: 1,
		},
		"ManufacturerFullAddress": {
			SearchTerms: []string{"full address"},
			CellRanges: []CellRange{
				{StartCell: "C18", EndCell: "C18"},
			},
			Offset: 1,
		},
		"ManufacturerCountry": {
			SearchTerms: []string{"Country"},
			CellRanges: []CellRange{
				{StartCell: "C19", EndCell: "C19"},
			},
			Offset: 1,
		},
		"ManufacturerCompanyNumber": {
			SearchTerms: []string{"company number"},
			CellRanges: []CellRange{
				{StartCell: "C20", EndCell: "C20"},
			},
			Offset: 1,
		},
		"CountryOfOrigin": {
			SearchTerms: []string{"country of origin"},
			CellRanges: []CellRange{
				{StartCell: "B21", EndCell: "D21"},
			},
			Offset: 3,
		},
		"CustomsTariffCode": {
			SearchTerms: []string{"customs tariff code"},
			CellRanges: []CellRange{
				{StartCell: "B22", EndCell: "D22"},
			},
			Offset: 3,
		},
		"ExportControlRegulated": {
			SearchTerms: []string{"export control regulations"},
			CellRanges: []CellRange{
				{StartCell: "B23", EndCell: "D23"},
			},
			DualColumnCheckBoxClf: true,
			DualColumnClfCriteria: DualClassificationCriteria{
				TYPE_1: ClassificationCriteria{
					Label:       "YES",
					SearchTerms: []string{"YES"},
					Offset:      3,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "NO",
					SearchTerms: []string{"No"},
					Offset:      4,
				},
			},
		},
		"PartClassification": {
			SearchTerms: []string{"classification of the part"},
			CellRanges: []CellRange{
				{StartCell: "B24", EndCell: "D24"},
			},
			TriColumnCheckBoxClf: true,
			TriColumnClfCriteria: TriClassificationCritera{
				TYPE_1: ClassificationCriteria{
					Label:       "DUAL",
					SearchTerms: []string{"DU", "DUAL"},
					Offset:      3,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "MILITARY",
					SearchTerms: []string{"MIL"},
					Offset:      3,
				},
				TYPE_3: ClassificationCriteria{
					Label:       "CIVIL",
					SearchTerms: []string{"CIVIL"},
					Offset:      5,
				},
			},
		},
		"ControlListClassificationNumber": {
			SearchTerms: []string{"control list classification number"},
			CellRanges: []CellRange{
				{StartCell: "B28", EndCell: "D28"},
			},
			Offset: 3,
		},
		"ThirdCountryControlledContent": {
			SearchTerms: []string{"third country controlled content"},
			CellRanges: []CellRange{
				{StartCell: "B29", EndCell: "D29"},
			},
			DualColumnCheckBoxClf: true,
			DualColumnClfCriteria: DualClassificationCriteria{
				TYPE_1: ClassificationCriteria{
					Label:       "YES",
					SearchTerms: []string{"YES"},
					Offset:      3,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "NO",
					SearchTerms: []string{"No"},
					Offset:      4,
				},
			},
		},
		"EndUserStatementRequired": {
			SearchTerms: []string{"end user statement will be required"},
			CellRanges: []CellRange{
				{StartCell: "B31", EndCell: "E31"},
			},
			DualColumnCheckBoxClf: true,
			DualColumnClfCriteria: DualClassificationCriteria{
				TYPE_1: ClassificationCriteria{
					Label:       "YES",
					SearchTerms: []string{"YES"},
					Offset:      4,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "NO",
					SearchTerms: []string{"No"},
					Offset:      4,
				},
			},
		},
		"ExportLicenceShipmentRequired": {
			SearchTerms: e.ReplaceCompanyNames([]string{"Export Licence for shipment to {companyName}"}),
			CellRanges: []CellRange{
				{StartCell: "B32", EndCell: "E32"},
			},
			DualColumnCheckBoxClf: true,
			DualColumnClfCriteria: DualClassificationCriteria{
				TYPE_1: ClassificationCriteria{
					Label:       "YES",
					SearchTerms: []string{"YES"},
					Offset:      4,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "NO",
					SearchTerms: []string{"NO"},
					Offset:      4,
				},
			},
		},
		"ExportLicenceEndUserRequired": {
			SearchTerms: e.ReplaceCompanyNames([]string{"Export Licence for shipment to {companyName} Specified End User"}),
			CellRanges: []CellRange{
				{StartCell: "B33", EndCell: "E33"},
			},
			TriColumnCheckBoxClf: true,
			TriColumnClfCriteria: TriClassificationCritera{
				TYPE_1: ClassificationCriteria{
					Label:       "YES",
					SearchTerms: []string{"YES"},
					Offset:      4,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "NO",
					SearchTerms: []string{"NO"},
					Offset:      4,
				},
				TYPE_3: ClassificationCriteria{
					Label:       "END USER NOT ADVISED TO SUPPLIER",
					SearchTerms: []string{"END USER NOT ADVISED TO SUPPLIER"},
					Offset:      4,
				},
			},
		},
		"AdditionalExportDocsRequired": {
			SearchTerms: []string{"Are other export documents required to be completed by"},
			CellRanges: []CellRange{
				{StartCell: "B34", EndCell: "E34"},
			},
			DualColumnCheckBoxClf: true,
			DualColumnClfCriteria: DualClassificationCriteria{
				TYPE_1: ClassificationCriteria{
					Label:       "YES",
					SearchTerms: []string{"YES"},
					Offset:      4,
				},
				TYPE_2: ClassificationCriteria{
					Label:       "NO",
					SearchTerms: []string{"No"},
					Offset:      4,
				},
			},
		},
		"TransferReexportConditions": {
			SearchTerms: []string{"additional is required to allow the product to be shipped"},
			CellRanges: []CellRange{
				{StartCell: "B35", EndCell: "E35"},
				{StartCell: "B36", EndCell: "E36"},
			},
			Offset: 4,
		},
		"RepresentativeName": {
			SearchTerms: []string{"name"},
			CellRanges: []CellRange{
				{StartCell: "B49", EndCell: "D49"},
				{StartCell: "B50", EndCell: "D50"},
			},
			Offset: 3,
		},
		"RepresentativePosition": {
			SearchTerms: []string{"position in the company"},
			CellRanges: []CellRange{
				{StartCell: "B50", EndCell: "D50"},
				{StartCell: "B51", EndCell: "D51"},
			},
			Offset: 3,
		},
		"RepresentativeSignature": {
			SearchTerms: []string{"Signature of Supplier"},
			CellRanges: []CellRange{
				{StartCell: "B51", EndCell: "D51"},
				{StartCell: "B52", EndCell: "D52"},
			},
			Offset:            3,
			BoolContainsImage: true,
			BoolClfContainsImage: BoolClassificationCriteria{
				Offset:      3,
				SearchTerms: []string{},
			},
		},
		"SupplierCompanySeal": {
			SearchTerms: []string{"SUPPLIER COMPANY SEAL", "company name"},
			CellRanges: []CellRange{
				{StartCell: "B52", EndCell: "D52"},
				{StartCell: "B53", EndCell: "D53"},
			},
			Offset: 3,
		},
		"SignatureDate": {
			SearchTerms: []string{"DATE", "(day/month/year)"},
			CellRanges: []CellRange{
				{StartCell: "B53", EndCell: "D53"},
				{StartCell: "B54", EndCell: "D54"},
			},
			Offset: 3,
		},
	}

	if err_product_sheet_search != nil {
		fmt.Println("%w", err_product_sheet_search)
	} else {
		e.Extraction.ProductDetails = &ProductDetails{
			SheetName: productSheetName,
		}
		e.extractDetails(e.Extraction.ProductDetails, productSheetName, productDetailsCriteria)
	}

	// Define column mappings with search terms
	columnMappings := []ColumnMapping{
		{
			FieldName:   "ItemNum",
			SearchTerms: []string{"Item"},
		},
		{
			FieldName:   "PartNumber",
			SearchTerms: []string{"part number"},
		},
		{
			FieldName:   "ComponentManufacturerPartNumber",
			SearchTerms: []string{"component manufacturer part number", "component manufacturer part-nr"},
		},
		{
			FieldName:   "PartDescription",
			SearchTerms: []string{"part description", "component description"},
		},
		{
			FieldName:   "ManufacturerOfComponent",
			SearchTerms: []string{"manufacturer of the component", "manufacturer of component"},
		},
		{
			FieldName:   "ExportRegulationCountry",
			SearchTerms: []string{"export regulations country"},
		},
		{
			FieldName:   "DualControlListClfNum",
			SearchTerms: []string{"Dual Use Item  - Control list classification number"},
		},
		{
			FieldName:   "MilitaryControlListClfNum",
			SearchTerms: []string{"Military Item - Control list classification number"},
		},
		{
			FieldName:   "IndicateLicenseApplication",
			SearchTerms: []string{"Indicate License Application Form/Type "},
		},
		{
			FieldName:   "TopLevelDeliverableItem",
			SearchTerms: []string{"Content of the top level deliverable item"},
		},
		{
			FieldName:   "USML_N",
			SearchTerms: []string{"usml n°", "usml"},
		},
		{
			FieldName:   "ECCN_N",
			SearchTerms: []string{"ECCN N°", "ECCN", "EAR 99"},
		},
		{
			FieldName:   "US_EA_CONTENT_RATIO",
			SearchTerms: []string{"Ratio of US EAR controlled content"},
		},
	}

	// Add controlled content extraction
	_, controlledContentSheetName, err := e.searchSheetName("controlled content")
	if err != nil {
		fmt.Println("Error finding controlled content sheet:", err)
	} else {
		e.Extraction.ControlledContent = e.extractControlledContent(controlledContentSheetName, columnMappings)
	}

	jsonBytes, err := json.Marshal(e.Extraction)
	if err != nil {
		fmt.Println("Error:", err)
		return *e.Extraction
	}
	fmt.Printf("Extraction with: %s \n", string(jsonBytes))

	return *e.Extraction
}

func (e *ExcelExtractor) Close() error {
	return e.file.Close()
}

///////////////////////
//// working on it ////
///////////////////////
