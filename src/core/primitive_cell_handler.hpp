#pragma once

#include <cstddef>
#include <string>

namespace xlsxcsv::internal {

// Internal worksheet fast path. It avoids constructing the public CellData
// variant for consumers that only need primitive worksheet values.
class PrimitiveCellHandler {
public:
    virtual ~PrimitiveCellHandler() = default;

    virtual void beginPrimitiveRow(
        int rowNumber, bool hidden, std::size_t columnReserveHint) = 0;
    virtual void addEmpty(int column) = 0;
    virtual void addBoolean(int column, bool value) = 0;
    virtual void addNumber(int column, double value, int styleIndex) = 0;
    virtual void addError(int column) = 0;
    virtual void addString(int column, std::string&& value) = 0;
    virtual void addSharedString(int column, std::size_t index) = 0;
    virtual void endPrimitiveRow() = 0;
};

// Internal opt-in control used by bounded consumers. SheetStreamReader checks
// it after each physical row so a consumer can finish without scanning the
// remainder of the worksheet XML.
class WorksheetRowControl {
public:
    virtual ~WorksheetRowControl() = default;
    virtual bool shouldParseRow(int rowNumber) const = 0;
    virtual bool shouldContinueParsing() const = 0;
};

} // namespace xlsxcsv::internal
