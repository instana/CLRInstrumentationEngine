#pragma once
#include <stdio.h>
#include <conio.h>
using namespace std;

namespace Outer
{
    inline void OuterSample()
    {
        Inner::InnerSample();
    }

    namespace Inner
    {
        void InnerSample();
    };
};