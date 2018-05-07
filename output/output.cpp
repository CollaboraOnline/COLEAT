/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */

#include "output.hpp"

int main(int, char**)
{
    Output::out() << "Hello there";
    {
        Output::out() << enter << "Should be one a new line, indented";
    }
    Output::out() << "Back at normal indent level\n";

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */
