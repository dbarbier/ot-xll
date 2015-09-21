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
#ifndef WIN32_LEAN_AND_MEAN
# define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>

#include <OT.hxx>
#include <vector>
#include "xll_helper_functions.h"

/***********************************************************************************
 OT_NORMAL_PDF()

 Purpose:

      This function takes 3 arguments and computes the normal distribution at a point.

 Parameters:

      LPXLOPER12      3 arguments : xl_mu, xl_sigma, xl_point
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
        return dialogError("(OT_NORMAL_PDF): Invalid conversion to xltypeNum for argument 'mu'", error);
    }

    // Coerce the standard deviation parameter
    //========================================
    if((error = xloper_to_num(xl_sigma, &sigma)) != -1)
    {
        return dialogError("(OT_NORMAL_PDF): Invalid conversion to xltypeNum for argument 'sigma'", error);
    }

    // Coerce the point parameter : to compute the PDF
    //================================================
    if((error = xloper_to_num(xl_point, &point)) != -1)
    {
        return dialogError("(OT_NORMAL_PDF): Invalid conversion to xltypeNum for argument 'point'", error);
    }

    double value;
    //Compute the PDF on point
    //========================
    try
    {
        OT::Normal distribution(mu, sigma);
        value = distribution.computePDF(point);
    }
    catch(OT::Exception & e)
    {
        return dialogError(e.what(), xlerrValue);
    }
    catch(std::exception & e)
    {
        return dialogError(e.what(), xlerrValue);
    }

    // xlbitDLLFree enables the DLL to release
    // any dynamically allocated memory
    // that was associated with the xloper
    LPXLOPER12 xResult = new XLOPER12();
    xResult->xltype = xltypeNum | xlbitDLLFree;
    xResult->val.num = value;

    return xResult;
}

/***********************************************************************************
 OT_NORMAL_PDF_ARRAY()

 Purpose:

      This function takes 3 arguments and computes the normal distribution at given points.

 Parameters:

      LPXLOPER12      3 arguments : xl_mu, xl_sigma, xl_points
                      (first two can be reference or value, last one is a reference)

 Returns:

      LPXLOPER12      the normal distribution at given points
                      or #VALUE! if there are
                      non-numerics in the supplied
                      argument.
*************************************************************************************/

LPXLOPER12 WINAPI
OT_NORMAL_PDF_ARRAY(LPXLOPER12 xl_mu, LPXLOPER12 xl_sigma, LPXLOPER12 xl_points)
{
    double mu, sigma;
    int error = -1;
    std::vector<double> points, pdf;

    // Coerce the mean parameter
    //==========================
    if((error = xloper_to_num(xl_mu, &mu)) != -1)
    {
        return dialogError("(OT_NORMAL_PDF_ARRAY): Invalid conversion to xltypeNum for argument 'mu'", error);
    }

    // Coerce the standard deviation parameter
    //========================================
    if((error = xloper_to_num(xl_sigma, &sigma)) != -1)
    {
        return dialogError("(OT_NORMAL_PDF_ARRAY): Invalid conversion to xltypeNum for argument 'sigma'", error);
    }

    // Coerce the points parameter : to compute the PDF
    //=================================================
    XLOPER12 cells;
    if((error = xloper_to_multi(xl_points, &cells)) != -1)
    {
        return dialogError("(OT_NORMAL_PDF_ARRAY): Invalid conversion to xltypeMulti for argument 'points'", error);
    }
    if (cells.val.array.columns != 1)
    {
        Excel12f(xlFree, 0, 1, (LPXLOPER12) &cells);
        return dialogError("(OT_NORMAL_PDF_ARRAY): Invalid active selection, there must be a single column", xlerrValue);
    }
    points.reserve(cells.val.array.rows);
    for(int i = 0; i < cells.val.array.rows; ++i)
    {
        points.push_back(cells.val.array.lparray[i].val.num);
    }
    // Delete cells to avoid leaks, this structure is no more needed
    Excel12f(xlFree, 0, 1, (LPXLOPER12) &cells);

    //Compute the PDF on points and store values in a vector
    //======================================================
    try
    {
        OT::Normal distribution(mu, sigma);
        OT::NumericalSample sampleInput(points.size(), 1);
        for(OT::UnsignedInteger i = 0; i < points.size(); ++i)
        {
            sampleInput[i][0] = points[i];
        }
        OT::NumericalSample samplePDF(distribution.computePDF(sampleInput));
        pdf.reserve(samplePDF.getSize());
        for(OT::UnsignedInteger i = 0; i < samplePDF.getSize(); ++i)
        {
            pdf.push_back(samplePDF[i][0]);
        }
    }
    catch(OT::Exception & e)
    {
        return dialogError(e.what(), xlerrValue);
    }
    catch(std::exception & e)
    {
        return dialogError(e.what(), xlerrValue);
    }
    if (points.size() != pdf.size())
    {
        return dialogError("(OT_NORMAL_PDF_ARRAY): Internal error", xlerrValue);
    }

    // Fill results
    //=============
    LPXLOPER12 xResult = new XLOPER12();

    // xlbitDLLFree enables the DLL to release
    // any dynamically allocated memory
    // that was associated with the xloper
    xResult->xltype = xltypeMulti | xlbitDLLFree;
    xResult->val.array.columns = 1;
    xResult->val.array.rows = pdf.size();
    xResult->val.array.lparray = new XLOPER12[xResult->val.array.rows*xResult->val.array.columns];

    LPXLOPER12 px = xResult->val.array.lparray;
    for(int i = 0; i < xResult->val.array.rows; ++i, ++px)
    {
        px->xltype = xltypeNum;
        px->val.num = pdf[i];
    }
    return xResult;
}

/***********************************************************************************
 OT_NORMAL_PDF_DRAW()

 Purpose:

      This function takes 2 arguments and computes a sample of the normal distribution,
      the number of points being given by the active cell selection

 Parameters:

      LPXLOPER12      2 arguments : xl_mu, xl_sigma
                      (can be reference or value)

 Returns:

      LPXLOPER12      the normal distribution at given points
                      or #VALUE! if there are
                      non-numerics in the supplied
                      argument.
*************************************************************************************/

LPXLOPER12 WINAPI
OT_NORMAL_PDF_DRAW(LPXLOPER12 xl_mu, LPXLOPER12 xl_sigma)
{
    double mu, sigma;
    int error = -1;
    std::vector<double> pdf;

    // Get the number of rows in current selection
    //============================================
    int nrValues = getNumberOfRows();
    if(nrValues <= 0)
    {
        return dialogError("(OT_NORMAL_PDF_DRAW): detection of cell selection failed", error);
    }
    int nrColumns = getNumberOfColumns();
    if(nrColumns == 0)
    {
        return dialogError("(OT_NORMAL_PDF_DRAW): detection of cell selection failed", error);
    }
    if(nrColumns != 2)
    {
        return dialogError("(OT_NORMAL_PDF_DRAW): wrong number of columns, two columns must be selected", error);
    }

    // Coerce the mean parameter
    //======================
    if((error = xloper_to_num(xl_mu, &mu)) != -1)
    {
        return dialogError("(OT_NORMAL_PDF_DRAW): Invalid conversion to xltypeNum for argument 'mu'", error);
    }

    // Coerce the standard deviation parameter
    //========================================
    if((error = xloper_to_num(xl_sigma, &sigma)) != -1)
    {
        return dialogError("(OT_NORMAL_PDF_DRAW): Invalid conversion to xltypeNum for argument 'sigma'", error);
    }

    //Compute the normal distribution and its PDF
    //===========================================
    try
    {
        OT::Normal distribution(mu, sigma);
        OT::NumericalSample samplePDF(distribution.drawPDF(nrValues).getDrawable(0).getData());
        pdf.reserve(2*samplePDF.getSize());
        for(OT::UnsignedInteger i = 0; i < samplePDF.getSize(); ++i)
        {
            pdf.push_back(samplePDF[i][0]);
            pdf.push_back(samplePDF[i][1]);
        }
    }
    catch(OT::Exception & e)
    {
        return dialogError(e.what(), xlerrValue);
    }
    catch(std::exception & e)
    {
        return dialogError(e.what(), xlerrValue);
    }
    if (2*nrValues != pdf.size())
    {
        return dialogError("(OT_NORMAL_PDF_DRAW): Internal error", xlerrValue);
    }

    // Fill results
    //=============
    LPXLOPER12 xResult = new XLOPER12();

    // xlbitDLLFree enables the DLL to release
    // any dynamically allocated memory
    // that was associated with the xloper
    xResult->xltype = xltypeMulti | xlbitDLLFree;
    xResult->val.array.columns = 2;
    xResult->val.array.rows = nrValues;
    xResult->val.array.lparray = new XLOPER12[xResult->val.array.rows*xResult->val.array.columns];

    LPXLOPER12 px = xResult->val.array.lparray;
    for(int i = 0; i < xResult->val.array.rows*xResult->val.array.columns; ++i, ++px)
    {
        px->xltype = xltypeNum;
        px->val.num = pdf[i];
    }
    return xResult;
}

/***********************************************************************************
 OT_NORMAL_PDF_DRAW_CMD()

 Purpose:

      This function takes 3 arguments and computes a sample of the normal distribution,
      the number of points being passed as an argument

 Parameters:

      LPXLOPER12      3 arguments : nrValues, xl_mu, xl_sigma
                      (can be reference or value)

 Returns:

      LPXLOPER12      the normal distribution at a point
                      or #VALUE! if there are
                      non-numerics in the supplied
                      argument.
*************************************************************************************/

LPXLOPER12 WINAPI
OT_NORMAL_PDF_DRAW_CMD(int nrValues, LPXLOPER12 xl_mu, LPXLOPER12 xl_sigma)
{
    double mu, sigma;
    int error = -1;
    std::vector<double> pdf;

    if(nrValues <= 0)
    {
        return dialogError("(OT_NORMAL_PDF_DRAW_CMD): first argument must be positive", error);
    }

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

    //Compute the PDF on point
    //========================
    try
    {
        OT::Normal distribution(mu, sigma);
        OT::NumericalSample samplePDF(distribution.drawPDF(nrValues).getDrawable(0).getData());
        pdf.reserve(2*samplePDF.getSize());
        for(OT::UnsignedInteger i = 0; i < samplePDF.getSize(); ++i)
        {
            pdf.push_back(samplePDF[i][0]);
            pdf.push_back(samplePDF[i][1]);
        }
    }
    catch(OT::Exception & e)
    {
        return dialogError(e.what(), xlerrValue);
    }
    catch(std::exception & e)
    {
        return dialogError(e.what(), xlerrValue);
    }
    if (2*nrValues != pdf.size())
    {
        return dialogError("(OT_NORMAL_PDF_DRAW_CMD): Internal error", xlerrValue);
    }   

    // Fill results
    //=============
    LPXLOPER12 xResult = new XLOPER12();

    // xlbitDLLFree enables the DLL to release
    // any dynamically allocated memory
    // that was associated with the xloper
    xResult->xltype = xltypeMulti | xlbitDLLFree;
    xResult->val.array.columns = 2;
    xResult->val.array.rows = nrValues;
    xResult->val.array.lparray = new XLOPER12[xResult->val.array.rows*xResult->val.array.columns];

    LPXLOPER12 px = xResult->val.array.lparray;
    for(int i = 0; i < xResult->val.array.rows * xResult->val.array.columns; ++i, ++px)
    {
        px->xltype = xltypeNum;
        px->val.num = pdf[i];
    }

    return xResult;
}

