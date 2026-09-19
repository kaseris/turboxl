#pragma once
#include "fixture_config.hpp"
#include <cstdlib>
#include <filesystem>
#include <initializer_list>
#include <stdexcept>
#include <string>

inline std::string fixtureQuote(const std::string& value) {
#ifdef _WIN32
    return "\"" + value + "\"";
#else
    std::string out = "'";
    for (char c : value) out += c == '\'' ? "'\\''" : std::string(1, c);
    return out + "'";
#endif
}

inline void createArchive(const std::filesystem::path& root,
                          const std::filesystem::path& output,
                          std::initializer_list<std::string> members = {"."}) {
    std::string command = fixtureQuote(FIXTURE_PYTHON) + " " + fixtureQuote(FIXTURE_SCRIPT)
        + " --archive " + fixtureQuote(root.string()) + " " + fixtureQuote(output.string());
    for (const auto& member : members) command += " " + fixtureQuote(member);
#ifdef _WIN32
    command = "\"" + command + "\"";
#endif
    if (std::system(command.c_str()) != 0 || !std::filesystem::exists(output))
        throw std::runtime_error("Could not create ZIP fixture: " + output.string());
}
