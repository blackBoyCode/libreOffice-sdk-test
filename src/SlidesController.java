import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawView;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.presentation.XPresentation;
import com.sun.star.presentation.XPresentation2;
import com.sun.star.presentation.XPresentationSupplier;
import com.sun.star.presentation.XSlideShowController;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;

/**
 * this class will control your impress slidesShow in editMode and PresentationMode
 * will use XDrawView(for edit ) XSlideShowController(for presentation mode)
 */
public class SlidesController {


    //for editMode
    private XDrawView xDrawView;

    //for fullScreenMode
    private XSlideShowController xSlideShowController;

    //add more unique functions to the original Xpresentation (Xpresentation != Xpresentation2  not the same thing)
    private XPresentation2 xPresentation2;

    //current PresentationDocument
    private XComponent currentPresentation;




    //will initialize the controller
    public SlidesController(XComponent currentPresentation){



        //FOR PRESENTATION MODE

        //assign presentation to a global variable in this class
        this.currentPresentation = currentPresentation;

        // presentation supplier instance for getting active presentation
        XPresentationSupplier xPresentationSupplier = UnoRuntime.queryInterface( XPresentationSupplier.class, currentPresentation );

        //assign active presentation
        XPresentation xPresentation = xPresentationSupplier.getPresentation();

        //enhance the xPresentation instance to access more methods (in this case the XSlideShowController)
        xPresentation2 = UnoRuntime.queryInterface(XPresentation2.class, xPresentation);

        //controller for full screen presentation
        xSlideShowController = xPresentation2.getController();

        //FOR EDIT MODE
        //need this for getting the controller for editMode
        XModel xModel = UnoRuntime.queryInterface(XModel.class, currentPresentation );

        //assign the controller from xModel method
        XController xController = xModel.getCurrentController();

        //get the DrawView to controlling
        xDrawView = UnoRuntime.queryInterface(XDrawView.class, xController);





    }


    public void nextSlide() throws WrappedTargetException, IndexOutOfBoundsException {

        //are we in fullScreenPresentation?
        if(xPresentation2.isRunning()){

            xSlideShowController.gotoNextSlide();


        }else{

            //TODO EXPERIMENTAL


            //TODO need to get current slide...

            //TODO then increment
            XDrawPage temp = getDrawPageForNextSlide(xDrawView.getCurrentPage());//BUG
            //TODO then setCurrentPage

            //use of xDrawViewer
            xDrawView.setCurrentPage(temp);





        }

    }


    public void previousSlide() throws WrappedTargetException, IndexOutOfBoundsException {

        //are we in fullScreenPresentation?
        if(xPresentation2.isRunning()){

            xSlideShowController.gotoPreviousSlide();

        }else{

            //use of xDrawViewer
            //TODO need to get current slide...

            //TODO then increment
            XDrawPage temp = getDrawPageForPreviousSlide(xDrawView.getCurrentPage());
            //TODO then setCurrentPage

            //use of xDrawViewer
            xDrawView.setCurrentPage(temp);



        }


    }


    // get the XDrawPage number for current slide
    private int getPageNumber(XDrawPage currentPage){

        int pageNumber = 0;

        XPropertySet xPropSet = UnoRuntime.queryInterface(XPropertySet.class, currentPage);
        try{
            pageNumber = AnyConverter.toInt(xPropSet.getPropertyValue("Number"));

            System.out.println("currentPageNumber: " + pageNumber);

        } catch (WrappedTargetException e) {
            e.printStackTrace();
        } catch (UnknownPropertyException e) {
            e.printStackTrace();
        }

        return pageNumber;
    }





    //TODO need a failsafe for NOT going indexOutOfBound from the presentation slides
    private XDrawPage getDrawPageForNextSlide(XDrawPage currentPage) throws WrappedTargetException, IndexOutOfBoundsException {
        //gives out an element i.e : if 3 slides returns 3
        int totalNumberOfSlides = PageHelper.getDrawPageCount(currentPresentation) ;
        System.out.println(totalNumberOfSlides);

        //current page for slide
        int currentPageNumber = getPageNumber(currentPage ) - 1 ;//start at zero "-1"



        if(currentPageNumber != totalNumberOfSlides - 1){

            return PageHelper.getDrawPageByIndex(currentPresentation, (currentPageNumber + 1));

        }else {

            return xDrawView.getCurrentPage();
        }


    }

    private XDrawPage getDrawPageForPreviousSlide(XDrawPage currentPage) throws WrappedTargetException, IndexOutOfBoundsException {



        //current page for slide
        int currentPageNumber = getPageNumber(currentPage) -1; //getPageNumber return an element a non zero value



        if(currentPageNumber != 0){ //first index is zero in most programming cases

            return PageHelper.getDrawPageByIndex(currentPresentation, (currentPageNumber - 1));

        }else {

            return xDrawView.getCurrentPage();
        }


    }



}
