#include "xlsxcsv/core.hpp"
#include <stdexcept>

namespace xlsxcsv::core {

class CsvEncoder {
public:
    CsvEncoder() = default;
    ~CsvEncoder() = default;
    
    // TODO: Implement in Phase 6
    std::string encode(const std::vector<std::vector<std::string>>& data) {
        throw std::runtime_error("CsvEncoder not implemented yet - coming in Phase 6");
    }
};

} // namespace xlsxcsv::core
