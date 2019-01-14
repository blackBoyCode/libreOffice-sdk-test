/* -*- Mode: Java; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*************************************************************************
 *
 *  The Contents of this file are made available subject to the terms of
 *  the BSD license.
 *
 *  Copyright 2000, 2010 Oracle and/or its affiliates.
 *  All rights reserved.
 *
 *  Redistribution and use in source and binary forms, with or without
 *  modification, are permitted provided that the following conditions
 *  are met:
 *  1. Redistributions of source code must retain the above copyright
 *     notice, this list of conditions and the following disclaimer.
 *  2. Redistributions in binary form must reproduce the above copyright
 *     notice, this list of conditions and the following disclaimer in the
 *     documentation and/or other materials provided with the distribution.
 *  3. Neither the name of Sun Microsystems, Inc. nor the names of its
 *     contributors may be used to endorse or promote products derived
 *     from this software without specific prior written permission.
 *
 *  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
 *  "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 *  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
 *  FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
 *  COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 *  INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 *  BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS
 *  OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 *  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR
 *  TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
 *  USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 *************************************************************************/

// __________ Imports __________

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.*;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.presentation.*;
import com.sun.star.uno.UnoRuntime;

import com.sun.star.frame.XController;
import com.sun.star.view.XSelectionChangeListener;
import com.sun.star.view.XSelectionSupplier;

import java.util.Scanner;


// __________ Implementation __________

// presentation demo

// This demo will demonstrate how to create a presentation using the Office API

// The first parameter describes the connection that is to use. If there is no parameter
// "uno:socket,host=localhost,port=2083;urp;StarOffice.ServiceManager" is used.


public class PresentationDemoSlides
{
    public static void main( String args[] )
    {
        XComponent xDrawDoc = null;
        try
        {
            // get the remote office context of a running office (a new office
            // instance is started if necessary)
            com.sun.star.uno.XComponentContext xOfficeContext = Helper.connect();


            // suppress Presentation Autopilot when opening the document
            // properties are the same as described for
            // com.sun.star.document.MediaDescriptor
            PropertyValue[] pPropValues = new PropertyValue[ 1 ];
            pPropValues[ 0 ] = new PropertyValue();
            pPropValues[ 0 ].Name = "Silent";
            pPropValues[ 0 ].Value = Boolean.TRUE;

            //TODO 1. Important create a new impress document
            //TODO need to select current active document for manipulation
            xDrawDoc = Helper.createDocument( xOfficeContext,
                "private:factory/simpress", "_blank", 0, pPropValues );





            XDrawPage    xPage;




            // create pages, so that three are available
            while ( PageHelper.getDrawPageCount( xDrawDoc ) < 3 )
                PageHelper.insertNewDrawPageByIndex( xDrawDoc, 0 );




            //TODO testing
            SlidesController slidesController = new SlidesController(xDrawDoc);


            // create two objects for the third page
            // clicking the first object lets the presentation jump
            // to page one by using ClickAction.FIRSTPAGE,
            // the second object lets the presentation jump to page two
            // by using a ClickAction.BOOKMARK;
            xPage = PageHelper.getDrawPageByIndex( xDrawDoc, 2 );


           // xShapePropSet.setPropertyValue(
//                ClickAction.NEXTPAGE
//                ClickAction.PREVPAGE



            /* start an endless presentation which is displayed in
               full-screen mode and placed on top */

            //those two lines are needed for XSlideShowController
            XPresentationSupplier xPresSupplier = UnoRuntime.queryInterface( XPresentationSupplier.class, xDrawDoc );
            XPresentation xPresentation = xPresSupplier.getPresentation();
//            XPropertySet xPresPropSet = UnoRuntime.queryInterface( XPropertySet.class, xPresentation );
//            xPresPropSet.setPropertyValue( "IsEndless", Boolean.TRUE );
//            xPresPropSet.setPropertyValue( "IsAlwaysOnTop", Boolean.TRUE );
//            xPresPropSet.setPropertyValue( "Pause", Integer.valueOf( 0 ) );
            //xPresentation.start();

         //   XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime.queryInterface(XDrawPagesSupplier.class, xDrawDoc);




            //TODO 2.important to for getting xController
          //  XModel  xModel = UnoRuntime.queryInterface(XModel.class, xDrawDoc );//need current or new impress document


            //TODO 3.assign xController with the XModel interface
            //XController xController = xModel.getCurrentController();


            //XSelectionSupplier xSelectionSupplier = UnoRuntime.queryInterface(XSelectionSupplier.class, xController);

            //XDrawView xDrawView = (XDrawView) xSelectionSupplier.getSelection();
            //xDrawView.setCurrentPage(xPage);

            //TODO CUSTOM CODE this works

            XPresentation2 xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(XPresentation2.class, xPresentation);

            PropertyValue[] aPresentationArgs = new PropertyValue[3];

            aPresentationArgs[0] = new PropertyValue();
            aPresentationArgs[0].Name = "IsAlwaysOnTop";
            aPresentationArgs[0].Value = Boolean.TRUE;

            aPresentationArgs[1] = new PropertyValue();
            aPresentationArgs[1].Name = "IsFullScreen";
            aPresentationArgs[1].Value = Boolean.FALSE;

            aPresentationArgs[2] = new PropertyValue();
            aPresentationArgs[2].Name = "IsAutomatic";
            aPresentationArgs[2].Value = Boolean.TRUE;



            //TODO I can also end the presentation....mmmm
            //xPresentation2.startWithArguments(aPresentationArgs);



            //TODO to verify if the presentation is running in full screen (we need XSlideShowController for this)
            //xPresentation2.isRunning();



            //TODO 4. need a XDrawView to set current active slide we pass in xController in the queryInterface
           // XDrawView xDrawView = UnoRuntime.queryInterface(XDrawView.class, xController );

            //XSelectionSupplier xSelectionSupplier = UnoRuntime.queryInterface(XSelectionSupplier.class, xDrawDoc);




            XDrawPage xDrawPage = PageHelper.getDrawPageByIndex(xDrawDoc, 2) ;


           // XController xController = UnoRuntime.queryInterface(XController.class, xDrawPage);

           // XDrawSubController xDrawSubController = UnoRuntime.queryInterface(XDrawSubController.class,xSelectionSupplier.getSelection());

            //xPresentation2.startWithArguments(aPresentationArgs);


            ///XSlideShowController xSlideShowController = xPresentation2.getController();


//            XSelectionSupplier xSelectionSupplier = UnoRuntime.queryInterface(XSelectionSupplier.class, xDrawDoc);






            while(true){

                Scanner in = new Scanner(System.in);
                int n = in.nextInt();

                if(n == 10){
                    slidesController.nextSlide();

                    //TODO when with start presentation in full screen
                    // xSlideShowController.gotoNextSlide();

                    //xDrawSubController.setCurrentPage(xDrawPage);

                    //TODO IMPORTANT the setCurrentPage() won't work in presentation mode.
                    //TODO 5. set current Page or slide in edit mode...?
                    //xDrawView.setCurrentPage(xDrawPage);

                    // ClickAction.NEXTPAGE;
                }else if(n== -10){
                    slidesController.previousSlide();
                }

            }



        }
        catch( Exception ex )
        {
            System.out.println( ex );
        }
        System.exit( 0 );









    }









    // this simple method applies the slide transition to a page
    public static void setSlideTransition( XDrawPage xPage,
        com.sun.star.presentation.FadeEffect eEffect,
            com.sun.star.presentation.AnimationSpeed eSpeed,
            int nChange,
                int nDuration )
    {
        // the following test is only sensible if you do not exactly know
        // what type of page xPage is, for this purpose it can been tested
        // if the com.sun.star.presentation.DrawPage service is supported
        XServiceInfo xInfo = UnoRuntime.queryInterface(
                XServiceInfo.class, xPage );
        if ( xInfo.supportsService( "com.sun.star.presentation.DrawPage" ) )
        {
            try
            {
                XPropertySet xPropSet = UnoRuntime.queryInterface( XPropertySet.class, xPage );
                xPropSet.setPropertyValue( "Effect",   eEffect );
                xPropSet.setPropertyValue( "Speed",    eSpeed );
                xPropSet.setPropertyValue( "Change",   Integer.valueOf( nChange ) );
                xPropSet.setPropertyValue( "Duration", Integer.valueOf( nDuration ) );
            }
            catch( Exception ex )
            {
            }
        }
    }
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
