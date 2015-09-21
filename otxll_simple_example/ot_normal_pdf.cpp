//                                               -*- C++ -*-
/**
 *  Copyright 2005-2015 Airbus-IMACS
 *
 *  This library is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU Lesser General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This library is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public
 *  along with this library.  If not, see <http://www.gnu.org/licenses/>.
 *
 */
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>

#include <OT.hxx>
#include "xll_helper_functions.h"

/***********************************************************************************
 OT_NORMAL_PDF()

 Purpose:

      This function takes 3 argument and computes the normal distribution at a point.

 Parameters:

      LPXLOPER12      3 argument : xl_mu, xl_sigma, xl_point
                      (can be reference or value)

 Returns:

      LPXLOPER12      the normal distribution at a point
                      or #VALUE! if there are
                      non-numerics in the supplied
                      argument.
*************************************************************************************/

LPXLOPER12 WINAPI
OT_NORMAL_PDF(LPXLOPER12 xl_mu, LPXLOPER12 xl_sigma, LPXLOPER12 xl_point)
{
    double mu, sigma, point;
    int error = -1;

    // Coerce the mean parameter
    //======================
    if((error = xloper_to_num(xl_mu, &mu)) != -1)
    {
        return dialogError("Invalid conversion to xltypeNum for argument 'mu'", error);
    }

    // Coerce the standard deviation parameter
    //========================================
    if((error = xloper_to_num(xl_sigma, &sigma)) != -1)
    {
        return dialogError("Invalid conversion to xltypeNum for argument 'sigma'", error);
    }

    // Coerce the point parameter : to compute the PDF
    //================================================
    if((error = xloper_to_num(xl_point, &point)) != -1)
    {
        return dialogError("Invalid conversion to xltypeNum for argument 'point'", error);
    }

    LPXLOPER12 xResult = new XLOPER12();
    //Compute the PDF on point
    //========================
    try
    {
        OT::Normal distribution(mu, sigma);
        // xlbitDLLFree enables the DLL to release
        // any dynamically allocated memory
        // that was associated with the xloper
        xResult->xltype = xltypeNum | xlbitDLLFree;
        xResult->val.num = distribution.computePDF(point);
    }
    catch(OT::Exception & e)
    {
        delete xResult;
        return dialogError(e.what(), xlerrValue);
    }
    catch(std::exception & e)
    {
        delete xResult;
        return dialogError(e.what(), xlerrValue);
    }

    return xResult;
}

