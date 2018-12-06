//
//  XLSReaderTests.swift
//  XLSReaderTests
//
//  Created by liaoguangwen on 2018/12/5.
//  Copyright © 2018 7owen. All rights reserved.
//

import XCTest
@testable import XLSReader

class XLSReaderTests: XCTestCase {

    override func setUp() {
        // Put setup code here. This method is called before the invocation of each test method in the class.
    }

    override func tearDown() {
        // Put teardown code here. This method is called after the invocation of each test method in the class.
    }

    func testExample() {
        // This is an example of a functional test case.
        // Use XCTAssert and related functions to verify your tests produce the correct results.
    }

    func testPerformanceExample() {
        XCTAssertEqual(xlsColToColStr(0), "A")
        XCTAssertEqual(xlsColToColStr(20), "U")
        XCTAssertEqual(xlsColToColStr(25), "Z")
        XCTAssertEqual(xlsColToColStr(26), "AA")
        XCTAssertEqual(xlsColToColStr(548), "UC")
        XCTAssertEqual(xlsColToColStr(619), "WV")
        XCTAssertEqual(xlsColToColStr(701), "ZZ")
        XCTAssertEqual(xlsColToColStr(702), "AAA")
        XCTAssertEqual(xlsColToColStr(1186), "ASQ")
        XCTAssertEqual(xlsColToColStr(1377), "AZZ")
        XCTAssertEqual(xlsColToColStr(1378), "BAA")
        XCTAssertEqual(xlsColToColStr(1420), "BBQ")
        XCTAssertEqual(xlsColToColStr(2053), "BZZ")
        XCTAssertEqual(xlsColToColStr(2054), "CAA")
        //XCTAssertEqual(xlsColToColStr(17575), "ZZZ")

        // This is an example of a performance test case.
        let path = Bundle.init(for: XLSReaderTests.self).path(forResource: "categories", ofType: "xls")!
        let workBook: XLSWorkBook! = XLSWorkBook(with: path)
        XCTAssertNotNil(workBook)
        XCTAssertEqual(workBook.numberOfSheets, 1)
        XCTAssertEqual(workBook.getSheetName(at: 0), "Sheet 1")

        let sheet: XLSWorkSheet! = workBook.getSheet(at: 0)
        XCTAssertNotNil(sheet)

        XCTAssertEqual(sheet.numberOfColsInSheet, 7)
        XCTAssertEqual(sheet.numberOfRowsInSheet, 15)

        let rows = sheet.numberOfRowsInSheet
        for row in 0..<rows {
            let cells = sheet.getCells(forRow: row)
            let rowString = cells.map { (cell) -> String in
                let suffix = " (\(cell.colStr)\(cell.row+1))"
                if cell.contentType != .blank {
                    let str = cell.content as? CustomStringConvertible ?? ""
                    return str.description + suffix
                } else {
                    return "<null>" + suffix
                }
                }.joined(separator: " | ")
            print(rowString)
        }

        self.measure {
            // Put the code you want to measure the time of here.
        }
    }

}
