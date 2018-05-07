/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora Office OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

// Experimental hacking, work in progress, not at all usable yet, and the whole idea will probably
// be re-worked a couple of times, if at all.

#ifndef INCLUDED_OUTPUT_HPP
#define INCLUDED_OUTPUT_HPP

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <cassert>
#include <cctype>
#include <codecvt>
#include <cstdio>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>

#pragma warning(pop)

class Output
{
private:
    static const int INDENTSTEP = 4;

    std::ostream& mrStream;
    unsigned mnIndentLevel;
    std::size_t mnLinePos;
    std::vector<std::string> mvBuffer;
    std::vector<unsigned> mvLineCount;

    Output(std::ostream& rStream)
        : mrStream(rStream)
        , mnIndentLevel(0)
        , mnLinePos(0)
    {
    }

public:
    class IncreaseIndent
    {
    public:
        IncreaseIndent() { Output::out().increaseIndent(); }

        ~IncreaseIndent() { Output::out().decreaseIndent(); }
    };

    Output& operator=(const Output&) = delete;

    static Output& out()
    {
        static Output snowflake(std::cout);

        return snowflake;
    }

    static std::ostream& stream() { return out().mrStream; }

    void increaseIndent() { mnIndentLevel++; }

    void decreaseIndent()
    {
        assert(mvBuffer.size() == mvLineCount.size());
        assert(mnIndentLevel > 0);

        if (mvLineCount.back() > 0)
        {
            mrStream << '\n';
            mnLinePos = 0;
        }
        mnIndentLevel--;
        mvBuffer.pop_back();
        mvLineCount.pop_back();
        if (mvBuffer.back().length() > 0)
        {
            mrStream << "..." << mvBuffer.back();
            mnLinePos += 3 + mvBuffer.back().length();
        }
    }

    Output& operator<<(const std::string& rString)
    {
        assert(mvBuffer.size() == mvLineCount.size());

        if (rString.length() > 0)
        {
            if (mnIndentLevel >= mvBuffer.size())
            {
                mrStream << "\n";
                mrStream << std::string(mnIndentLevel * INDENTSTEP, ' ');
            }
            std::string::size_type nSegmentBegin = 0;
            std::string::size_type nNewline;
            if (mvBuffer.size() <= mnIndentLevel)
            {
                mvBuffer.push_back("");
                mvLineCount.push_back(0);
            }
            while ((nNewline = rString.find('\n', nSegmentBegin)) != std::string::npos)
            {
                mvLineCount.back()++;
                mrStream << rString.substr(nSegmentBegin, nNewline - nSegmentBegin);
                mvBuffer[mnIndentLevel] = "";
                mrStream << std::string(mnIndentLevel * INDENTSTEP, ' ');
                mnLinePos = mnIndentLevel * INDENTSTEP;
                nSegmentBegin = nNewline + 1;
            }
            if (nSegmentBegin < rString.length())
            {
                mrStream << rString.substr(nSegmentBegin);
                if (mnLinePos == 0)
                    mvLineCount.back()++;
                mnLinePos += rString.length() - nSegmentBegin;
                mvBuffer[mnIndentLevel] = rString.substr(nSegmentBegin);
            }
        }
        return *this;
    }

    Output& operator<<(std::ostream& (*pManipulator)(std::ostream&))
    {
        mrStream << pManipulator;
        return *this;
    }

    Output& endl(Output&)
    {
        *this << "\n";
        mrStream << std::flush;
        return *this;
    }
};

std::ostream& enter(std::ostream& os)
{
    Output::out().increaseIndent();
    return os;
}

#endif // INCLUDED_OUTPUT_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */
