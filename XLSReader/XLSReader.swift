//
//  XLSReader.swift
//  XLSReader
//
//  Created by liaoguangwen on 2018/12/5.
//  Copyright © 2018 7owen. All rights reserved.
//

import Foundation
import libxls

public class XLSWorkBook {
    private var xlsWorkBook: xlsWorkBook

    public init?(with path: String) {
        guard let pxlsWorkBook = xls_open(path, "UTF-8") else {
            return nil
        }
        self.xlsWorkBook = pxlsWorkBook.pointee
        xls_parseWorkBook(pxlsWorkBook)
    }

    public var numberOfSheets: Int {
        return Int(xlsWorkBook.sheets.count)
    }

    public func isSheetVisible(at num: Int) -> Bool {
        let sheet = self.getSTSheetData(at: num)
        return sheet?.visibility != 0
    }

    public func getSheetName(at num: Int) -> String? {
        let sheet = self.getSTSheetData(at: num)
        return sheet != nil ? String(cString: sheet!.name) : nil
    }

    public func getSheet(at num: Int) -> XLSWorkSheet? {
        guard num < numberOfSheets else {
            return nil
        }

        if let sheet = xls_getWorkSheet(&xlsWorkBook, Int32(num)) {
            return XLSWorkSheet(with: sheet.pointee)
        } else {
            return nil
        }
    }

    private func getSTSheetData(at num: Int) -> st_sheet_data? {
        guard num < numberOfSheets else {
            return nil
        }
        return xlsWorkBook.sheets.sheet[num]
    }
}

public class XLSWorkSheet {
    private var xlsWorkSheet: xlsWorkSheet

    init(with sheet: xlsWorkSheet) {
        xlsWorkSheet = sheet
        xls_parseWorkSheet(&xlsWorkSheet)
    }

    public var numberOfRowsInSheet: Int {
        return Int(xlsWorkSheet.rows.lastrow + 1)
    }

    public var numberOfColsInSheet: Int {
        return Int(xlsWorkSheet.rows.lastcol + 1)
    }

    public func cellsIterator() -> Iterator {
        return CellsIterator(xlsWorkSheet)
    }

    public func getCells(forRow row: Int) -> [XLSCell] {
        guard row < self.numberOfRowsInSheet else {
            return []
        }
        var cells = [XLSCell]()
        let rowP = self.xlsWorkSheet.rows.row[row]
        for i in 0..<self.numberOfColsInSheet {
            let cell = rowP.cells.cell[i]
            cells.append(createCell(for: cell))
        }
        return cells
    }
}

public struct XLSCell {
    public enum ContentType: Int {
        case blank
        case string
        case float
        case bool
        case error
        case unknown
    }
    public let id: Int
    public let contentType: ContentType
    public let row: Int
    public let col: Int
    public let colStr: String
    public let content: Any?

    init(id: Int, contentType: ContentType, row: Int, col: Int, content: Any? = nil) {
        self.id = id
        self.contentType = contentType
        self.row = row
        self.col = col
        self.content = content
        self.colStr = xlsColToColStr(col)
    }
}

public protocol Iterator {
    func next() -> XLSCell?
}

class CellsIterator: Iterator {
    private let xlsWorkSheet: xlsWorkSheet
    private let numRows: UInt16
    private let numCols: UInt16
    private var row: Int = 0
    private var col: Int = 0

    init(_ xlsWorkSheet: xlsWorkSheet) {
        self.xlsWorkSheet = xlsWorkSheet
        numRows = xlsWorkSheet.rows.lastrow + 1
        numCols = xlsWorkSheet.rows.lastcol + 1
    }

    public func next() -> XLSCell? {
        guard self.row < self.numRows else {
            return nil
        }
        let rowP = self.xlsWorkSheet.rows.row[self.row]
        let cell = rowP.cells.cell[self.col]
        self.col += 1
        if self.col >= self.numCols {
            self.col = 0
            self.row += 1
        }
        return createCell(for: cell)
    }
}

private func createCell(for xlsCell: xlsCell) -> XLSCell {
    let col = xlsCell.col
    let row = xlsCell.row
    let type: XLSCell.ContentType
    var val: Any?

    switch Int32(xlsCell.id) {
    case XLS_RECORD_FORMULA:
        if xlsCell.l == 0 {
            type = .float
            val = Float(xlsCell.d)
        } else {
            switch xlsCell.str.hashValue {
            case "bool".hashValue:
                type = .bool
                val = Int(xlsCell.d) != 0
            case "error".hashValue:
                type = .error
                val = Int(xlsCell.d)
            default:
                type = .string
            }
        }
    case XLS_RECORD_LABELSST, XLS_RECORD_LABEL:
        type = .string
    case XLS_RECORD_NUMBER, XLS_RECORD_RK:
        type = .float
        val = Float(xlsCell.d)
    case XLS_RECORD_BLANK:
        type = .blank
    default:
        type = .unknown
    }
    if val == nil && xlsCell.str != nil {
        val = String(cString: xlsCell.str)
    }

    return XLSCell(id: Int(xlsCell.id), contentType: type, row: Int(row), col: Int(col), content: val)
}


public func xlsColToColStr(_ col: Int) -> String {
    let asciiA = Int("A".utf8.first!)
    if col < 26 {
        return String(UnicodeScalar(asciiA + col)!)
    } else if col/26 <= 26 {
        return String(UnicodeScalar(asciiA + col/26-1)!) + String(UnicodeScalar(asciiA + col%26)!)
    } else {
        return String(UnicodeScalar(asciiA + (col/26-1)/26-1)!) + String(UnicodeScalar(asciiA + (col/26-1)%26)!) + String(UnicodeScalar(asciiA + col%26)!)
    }
}
