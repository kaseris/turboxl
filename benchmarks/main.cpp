#include <benchmark/benchmark.h>
#include "xlsxcsv/core.hpp"

namespace {

xlsxcsv::core::RowData mixedRow() {
    using namespace xlsxcsv::core;
    RowData row;
    row.rowNumber = 1;
    for (int column = 1; column <= 8; ++column) {
        CellData cell;
        cell.coordinate = {1, column};
        if (column == 1 || column == 6) {
            cell.type = CellType::InlineString;
            cell.value = std::string("value, with \"quotes\"");
        } else if (column == 4) {
            cell.type = CellType::Boolean;
            cell.value = true;
        } else {
            cell.type = CellType::Number;
            cell.value = 12345.678901;
        }
        row.cells.push_back(std::move(cell));
    }
    return row;
}

void BM_MixedRowEncoding(benchmark::State& state) {
    const auto row = mixedRow();
    for (auto _ : state) {
        xlsxcsv::core::CsvRowCollector collector;
        for (int index = 0; index < 1000; ++index) collector.handleRow(row);
        benchmark::DoNotOptimize(collector.takeCsvString());
    }
    state.SetItemsProcessed(state.iterations() * 8000);
}

void BM_SparseRowEncoding(benchmark::State& state) {
    xlsxcsv::core::RowData row;
    row.rowNumber = 1;
    xlsxcsv::core::CellData cell;
    cell.coordinate = {1, 256};
    cell.type = xlsxcsv::core::CellType::Number;
    cell.value = 42.0;
    row.cells.push_back(cell);
    for (auto _ : state) {
        xlsxcsv::core::CsvRowCollector collector;
        collector.handleRow(row);
        benchmark::DoNotOptimize(collector.takeCsvString());
    }
}

BENCHMARK(BM_MixedRowEncoding);
BENCHMARK(BM_SparseRowEncoding);

} // namespace

BENCHMARK_MAIN();
