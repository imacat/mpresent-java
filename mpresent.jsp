<%
/*
 * mpresent.jsp
 *
 * Created on 2010-07-13 (Original CGI 2010-07-07)
 * 
 * Copyright (c) 2010 imacat
 *
 * version 0.01
 */
%><%@ page
contentType="text/html; charset=utf-8"

import="com.sun.star.beans.PropertyValue"
import="com.sun.star.bridge.XUnoUrlResolver"
import="com.sun.star.comp.helper.Bootstrap"
import="com.sun.star.frame.XDesktop"
import="com.sun.star.lang.XComponent"
import="com.sun.star.lang.XMultiComponentFactory"
import="com.sun.star.uno.UnoRuntime"
import="com.sun.star.uno.XComponentContext"
import="com.sun.star.ucb.XFileIdentifierConverter"

import="com.sun.star.container.XEnumeration"
import="com.sun.star.container.XEnumerationAccess"
import="com.sun.star.frame.XModel"
import="com.sun.star.lang.XServiceInfo"

import="com.sun.star.drawing.XDrawPages"
import="com.sun.star.drawing.XDrawPagesSupplier"

import="com.sun.star.frame.XComponentLoader"
import="com.sun.star.presentation.XPresentation"
import="com.sun.star.presentation.XPresentation2"
import="com.sun.star.presentation.XPresentationSupplier"
import="com.sun.star.presentation.XSlideShowController"

import="java.util.Hashtable"
import="java.util.Vector"
%><%!
    /** The servlet request */
    private HttpServletRequest r = null;
    
    /** The desktop service. */
    private Object desktop = null;
    
    /** The bootstrap context. */
    private XComponentContext bootstrapContext = null;
    
    /** The registry service manager. */
    private XMultiComponentFactory serviceManager = null;
    
    /** The document. */
    private XComponent doc = null;
    
    /** The action type. */
    private ActionType action = null;
    
    /** The file. */
    private String file = null;
    
    /** The current slide index. */
    private int currentSlide = -1;
    
    /** The goto slide index. */
    private int slide = -1;
    
    /** The slide indeces of all the opened presentations. */
    private Hashtable<String,Integer> savedSlideIndeces
        = new Hashtable<String,Integer>();
    
    /** The presentation document files. */
    private Vector<PresentationFile> allFiles
        = new Vector<PresentationFile>();
    
    /**
     * Parses the arguments.
     *
     */
    private void parseArguments() {
        this.action = null;
        this.file = null;
        this.currentSlide = -1;
        this.slide = -1;
        if (this.allFiles.size() == 0) {
            this.allFiles.add(new PresentationFile(
                "/home/imacat/ooomagic.odp",
                "OpenOffice.org的UNO魔術"));
            this.allFiles.add(new PresentationFile(
                "/home/imacat/（一）初級篇.odp",
                "Perl Regular Expression初級篇"));
            this.allFiles.add(new PresentationFile(
                "/home/imacat/（二）進階篇.odp",
                "Perl Regular Expression進階篇"));
            this.allFiles.add(new PresentationFile(
                "/home/imacat/ooomagic-real.odp",
                "真•OpenOffice.org的UNO魔術"));
            this.allFiles.add(new PresentationFile(
                "/home/imacat/ooomagic-real-real.odp",
                "真•真•OpenOffice.org的UNO魔術"));
        }
        
        try {
            this.r.setCharacterEncoding("UTF-8");
        } catch (java.io.UnsupportedEncodingException e) {
            throw new java.lang.IllegalArgumentException(e);
        }
        
        if (this.r.getMethod().equals("POST")) {
            if (this.r.getParameter("start") != null) {
                this.action = ActionType.START;
            } else if (this.r.getParameter("next") != null) {
                this.action = ActionType.NEXT;
            } else if (this.r.getParameter("previous") != null) {
                this.action = ActionType.PREVIOUS;
            } else if (this.r.getParameter("stop") != null) {
                this.action = ActionType.STOP;
            } else if (this.r.getParameter("goto") != null) {
                this.action = ActionType.GOTO;
            } else if (this.r.getParameter("close") != null) {
                this.action = ActionType.CLOSE;
            }
            this.file = this.r.getParameter("file");
            if (this.r.getParameter("slide") != null && !this.r.getParameter("slide").equals("")) {
                this.slide = Integer.parseInt(this.r.getParameter("slide")) - 1;
            }
        }
        
        return;
    }
    
    /**
     * Shows the HTML presentation controller page.
     *
     * @return the HTML of the page
     */
    private String showHTML() {
        String html = null;
        String script = this.r.getServletPath();
        int pos = 0;
        
        pos = script.lastIndexOf('/');
        script = script.substring(pos + 1);
        
        html = "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n"
            + "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.1//EN\"\n"
            + "    \"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd\">\n"
            + "<html xmlns=\"http://www.w3.org/1999/xhtml\" xml:lang=\"zh-tw\">\n"
            + "<head>\n"
            + "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />\n"
            + "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.5\" />\n"
            + "<title>簡報遙控</title>\n"
            + "</head>\n"
            + "<body>\n";
        
        if (this.doc != null) {
            XDrawPagesSupplier xDrawPagesSupplier = null;
            int count = -1;
            
            xDrawPagesSupplier = (XDrawPagesSupplier)
                UnoRuntime.queryInterface(
                XDrawPagesSupplier.class, this.doc);
            count = xDrawPagesSupplier.getDrawPages().getCount();
            
            html += "<form action=\"" + script + "\" method=\"POST\">\n"
                + "<div>\n"
                + "<input type=\"hidden\" name=\"file\" value=\""
                    + this.file + "\" />\n"
                + "<input type=\"submit\" name=\"stop\" value=\"■\" />\n"
                + "<input type=\"submit\" name=\"previous\" value=\"◀◀\" />\n"
                + "<input type=\"submit\" name=\"start\" value=\"▶\" />\n"
                + "<input type=\"submit\" name=\"next\" value=\"▶▶\" />\n"
                + "</div>\n"
                + "<div>&nbsp;</div>\n";
            
            html += "<div>\n"
                + "<select name=\"slide\">\n";
            for (int i = 0; i < count; i++) {
                html += "    <option value=\"" + (i + 1) + "\""
                    + (i == this.currentSlide? " selected=\"selected\"": "")
                    + ">" + (i + 1) + "</option>\n";
            }
            html += "</select>\n"
                + "<input type=\"submit\" name=\"goto\" value=\"Go\" />\n"
                + "<input type=\"submit\" name=\"close\" value=\"▲ \" />\n"
                + "</div>\n"
                + "</form>\n";
        }
        
        html += "<form action=\"" + script + "\" method=\"POST\">\n"
            + "<div><select name=\"file\">\n";
        for (int i = 0; i < this.allFiles.size(); i++) {
            PresentationFile thisFile = this.allFiles.get(i);
            
            html += "    <option"
                + " value=\"" + thisFile.getFile() + "\""
                + (this.file != null && this.file.equals(thisFile.getFile())?
                    " selected=\"selected\"": "")
                + ">" + thisFile.getTitle() + "</option>\n";
        }
        html += "</select></div>\n"
            + "<input type=\"submit\" name=\"start\" value=\"▶\" />\n"
            + "</div>\n"
            + "</form>\n"
            + "</body>\n"
            + "</html>";
        
        return html;
    }
    
    /**
     * Connects to the OpenOffice.org process.
     *
     * @param host specifies the host name
     * @param port specifies the port number
     * @throws com.sun.star.comp.helper.BootstrapException if fails to
     *         create the initial component context
     * @throws com.sun.star.connection.NoConnectException if no one
     *         is accepting on the specified resource. 
     * @throws com.sun.star.connection.ConnectionSetupException if
     *         it is not possible to accept on a local resource
     */
    private void connect(String host, int port)
            throws com.sun.star.comp.helper.BootstrapException,
                com.sun.star.connection.NoConnectException,
                com.sun.star.connection.ConnectionSetupException {
        boolean doConnect = false;
        XComponentContext localContext;
        XMultiComponentFactory localServiceManager = null;
        Object unoUrlResolver = null;
        XUnoUrlResolver xUnoUrlResolver = null;
        Object bootstrapContext = null;
        
        if (this.desktop == null) {
            doConnect = true;
        } else {
            try {
                UnoRuntime.queryInterface(XDesktop.class, this.desktop);
            } catch (com.sun.star.lang.DisposedException e) {
                doConnect = true;
            }
        }
        if (!doConnect) {
            return;
        }
        
        // Obtain the local context
        try {
            localContext = Bootstrap.createInitialComponentContext(null);
        } catch (java.lang.Exception e) {
            throw new com.sun.star.comp.helper.BootstrapException(e);
        }
        if (localContext == null) {
            throw new com.sun.star.comp.helper.BootstrapException(
                    "Failed to obtain the local OpenOffice.org component context.");
        }
        
        // Obtain the local service manager
        localServiceManager = localContext.getServiceManager();
        
        // Obtain the URL resolver
        try {
            unoUrlResolver = localServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", localContext);
        } catch (com.sun.star.uno.Exception e) {
            throw new java.lang.UnsupportedOperationException(e);
        }
        xUnoUrlResolver = (XUnoUrlResolver) UnoRuntime.queryInterface(
            XUnoUrlResolver.class, unoUrlResolver);
        
        // Obtain the context
        try {
            bootstrapContext = xUnoUrlResolver.resolve(String.format(
                "uno:socket,host=%s,port=%d;urp;StarOffice.ComponentContext",
                host, port));
        } catch (com.sun.star.connection.NoConnectException e) {
            throw e;
        } catch (com.sun.star.connection.ConnectionSetupException e) {
            throw e;
        } catch (com.sun.star.lang.IllegalArgumentException e) {
            throw new java.lang.IllegalArgumentException(e);
        }
        if (bootstrapContext == null) {
            throw new java.lang.UnsupportedOperationException(
                String.format("Failed to connect to OpenOffice.org at %s:%d.",
                    host, port));
        }
        this.bootstrapContext = (XComponentContext) UnoRuntime.queryInterface(
            XComponentContext.class, bootstrapContext);
        
        // Obtain the service manager
        this.serviceManager = this.bootstrapContext.getServiceManager();
        
        // Obtain the desktop service
        try {
            this.desktop = this.serviceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", this.bootstrapContext);
        } catch (com.sun.star.uno.Exception e) {
            throw new java.lang.UnsupportedOperationException(e);
        }
    }
    
    /**
     * Finds the already-opened document.
     *
     */
    private void findDocument() {
        Object fileContentProvider = null;
        XFileIdentifierConverter xFileIdentifierConverter = null;
        String url = null;
        XDesktop xDesktop = null;
        XEnumerationAccess xEnumerationAccess = null;
        XEnumeration xEnumeration = null;
        boolean isRunning = false;
        
        try {
            fileContentProvider
                = this.serviceManager.createInstanceWithContext(
                    "com.sun.star.ucb.FileContentProvider",
                    this.bootstrapContext);
        } catch (com.sun.star.uno.Exception e) {
            throw new java.lang.UnsupportedOperationException(e);
        }
        xFileIdentifierConverter = (XFileIdentifierConverter)
            UnoRuntime.queryInterface(
            XFileIdentifierConverter.class, fileContentProvider);
        url = xFileIdentifierConverter.getFileURLFromSystemPath(
            "", this.file);
        
        this.doc = null;
        
        // Obtain all the components
        xDesktop = (XDesktop) UnoRuntime.queryInterface(
            XDesktop.class, this.desktop);
        xEnumerationAccess = xDesktop.getComponents();
        xEnumeration = xEnumerationAccess.createEnumeration();
        while (xEnumeration.hasMoreElements()) {
            Object component = null;
            XServiceInfo xServiceInfo = null;
            XPresentationSupplier xPresentationSupplier = null;
            XPresentation xPresentation = null;
            XPresentation2 xPresentation2 = null;
            XModel xModel = null;
            boolean isThisDocument = false;
            
            try {
                component = xEnumeration.nextElement();
            } catch (com.sun.star.container.NoSuchElementException e) {
                throw new java.util.NoSuchElementException(e.getMessage());
            } catch (com.sun.star.lang.WrappedTargetException e) {
                throw new java.lang.RuntimeException(e);
            }
            
            xServiceInfo = (XServiceInfo) UnoRuntime.queryInterface(
                XServiceInfo.class, component);
            if (xServiceInfo == null) {
                continue;
            }
            if (!xServiceInfo.supportsService(
                    "com.sun.star.presentation.PresentationDocument")) {
                continue;
            }
            
            xModel = (XModel) UnoRuntime.queryInterface(
                XModel.class, component);
            if (url.equals(xModel.getURL())) {
                this.doc = (XComponent) UnoRuntime.queryInterface(
                    XComponent.class, component);
                isThisDocument = true;
            }
            
            xPresentationSupplier = (XPresentationSupplier)
                UnoRuntime.queryInterface(
                XPresentationSupplier.class, component);
            xPresentation = xPresentationSupplier.getPresentation();
            xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
                XPresentation2.class, xPresentation);
            if (!isThisDocument && xPresentation2.isRunning()) {
                XSlideShowController xSlideShowController
                    = xPresentation2.getController();
                
                this.savedSlideIndeces.put(xModel.getURL(),
                    xSlideShowController.getCurrentSlideIndex());
                xPresentation2.end();
            }
        }
    }
    
    /**
     * Opens the presentation document.
     *
     * @throws java.io.IOException when file couldn't be found or was
               corrupt
     */
    private void open()
            throws java.io.IOException {
        Object fileContentProvider = null;
        XFileIdentifierConverter xFileIdentifierConverter = null;
        String url = null;
        XComponentLoader xComponentLoader = null;
        PropertyValue props[] = new PropertyValue[0];
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        if (this.doc == null) {
            try {
                fileContentProvider
                    = this.serviceManager.createInstanceWithContext(
                        "com.sun.star.ucb.FileContentProvider",
                        this.bootstrapContext);
            } catch (com.sun.star.uno.Exception e) {
                throw new java.lang.UnsupportedOperationException(e);
            }
            xFileIdentifierConverter = (XFileIdentifierConverter)
                UnoRuntime.queryInterface(
                XFileIdentifierConverter.class, fileContentProvider);
            
            // Open the file
            url = xFileIdentifierConverter.getFileURLFromSystemPath(
                "", this.file);
            xComponentLoader = (XComponentLoader) UnoRuntime.queryInterface(
                XComponentLoader.class, this.desktop);
            try {
                this.doc = xComponentLoader.loadComponentFromURL(
                    url, "_default", 0, props);
            } catch (com.sun.star.io.IOException e) {
                throw new java.io.IOException(e);
            } catch (com.sun.star.lang.IllegalArgumentException e) {
                throw new java.lang.IllegalArgumentException(e);
            }
            
            if (this.savedSlideIndeces.containsKey(url)) {
                this.savedSlideIndeces.remove(url);
            }
        }
        
        // Obtain the presentation
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.doc);
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        
        // Obtain the current slide
        if (xPresentation2.isRunning()) {
            xSlideShowController = xPresentation2.getController();
            this.currentSlide = xSlideShowController.getCurrentSlideIndex();
        
        // Start the presentation if not started yet
        } else {
            if (this.action != ActionType.STOP) {
                XModel xModel = null;
                
                xPresentation2.start();
                while (!xPresentation2.isRunning()) { }
                xSlideShowController = xPresentation2.getController();
                
                xModel = (XModel) UnoRuntime.queryInterface(
                    XModel.class, this.doc);
                url = xModel.getURL();
                if (this.savedSlideIndeces.containsKey(url)) {
                    this.currentSlide = this.savedSlideIndeces.get(url);
                } else {
                    this.currentSlide = -1;
                }
                
                // Go to the current slide
                if (this.currentSlide == -1) {
                    if (xSlideShowController.getCurrentSlideIndex() != 0) {
                        xSlideShowController.gotoSlideIndex(0);
                    }
                } else {
                    if (    xSlideShowController.getCurrentSlideIndex()
                            != this.currentSlide) {
                        xSlideShowController.gotoSlideIndex(this.currentSlide);
                    }
                }
                
                // Redraw the screen for sometimes the screen is black
                xSlideShowController.gotoNextSlide();
                try {
                    Thread.currentThread().sleep(50);
                } catch (java.lang.InterruptedException e) {
                }
                xSlideShowController.gotoPreviousSlide();
            }
        }
        
        return;
    }
    
    /**
     * Starts the presentation.
     *
     */
    private void start() {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.doc);
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        xSlideShowController = xPresentation2.getController();
        
        this.currentSlide = xSlideShowController.getCurrentSlideIndex();
        return;
    }
    
    /**
     * Goes to the next slide.
     *
     */
    private void gotoNextSlide() {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.doc);
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        xSlideShowController = xPresentation2.getController();
        
        xSlideShowController.gotoNextSlide();
        this.currentSlide = xSlideShowController.getCurrentSlideIndex();
        return;
    }
    
    /**
     * Goes to the previous slide.
     *
     */
    private void gotoPreviousSlide() {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.doc);
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        xSlideShowController = xPresentation2.getController();
        
        xSlideShowController.gotoPreviousSlide();
        this.currentSlide = xSlideShowController.getCurrentSlideIndex();
        return;
    }
    
    /**
     * Stops the presentation
     *
     */
    private void stop() {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.doc);
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        
        if (xPresentation2.isRunning()) {
            XSlideShowController xSlideShowController
                = xPresentation2.getController();
            XModel xModel = (XModel) UnoRuntime.queryInterface(
                XModel.class, this.doc);
            
            this.savedSlideIndeces.put(xModel.getURL(),
                xSlideShowController.getCurrentSlideIndex());
            xPresentation.end();
        }
        return;
    }
    
    /**
     * Goes to a specific slide.
     *
     */
    private void gotoSlide() {
        XPresentationSupplier xPresentationSupplier = null;
        XPresentation xPresentation = null;
        XPresentation2 xPresentation2 = null;
        XSlideShowController xSlideShowController = null;
        
        xPresentationSupplier = (XPresentationSupplier)
            UnoRuntime.queryInterface(
            XPresentationSupplier.class, this.doc);
        xPresentation = xPresentationSupplier.getPresentation();
        xPresentation2 = (XPresentation2) UnoRuntime.queryInterface(
            XPresentation2.class, xPresentation);
        xSlideShowController = xPresentation2.getController();
        
        xSlideShowController.gotoSlideIndex(this.slide);
        this.currentSlide = xSlideShowController.getCurrentSlideIndex();
        return;
    }
    
    /**
     * Closes the presentation document.
     *
     */
    private void close() {
        this.doc.dispose();
        this.doc = null;
        return;
    }
    
    /**
     * The action type.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.01
     */
    private enum ActionType {
        
        /** Starts the presentation. */
        START,
        
        /** Goes to the next slide. */
        NEXT,
        
        /** Goes to the previous slide. */
        PREVIOUS,
        
        /** Stop the presentation. */
        STOP,
        
        /** Goes to a specific slide. */
        GOTO,
        
        /** Closes the presentation document. */
        CLOSE;
    }
    
    /**
     * The presentation file.
     *
     * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
     * @version 0.01
     */
    private class PresentationFile {
        
        /** The document file. */
        private String file = null;
        
        /** The document title. */
        private String title = null;
        
        /**
         * Creates a new instance of PresentationFile.
         *
         * @param file the document file
         * @param title the document title
         */
        private PresentationFile(String file, String title) {
            this.file = file;
            this.title = title;
        }
        
        /**
         * Returns the document file.
         *
         * @return the document file
         */
        private String getFile() {
            return this.file;
        }
        
        /**
         * Returns the document title.
         *
         * @return the document title
         */
        private String getTitle() {
            return this.title;
        }
    }
%><%
    this.r = request;
    this.parseArguments();
    if (this.action != null) {
        this.connect("localhost", 2002);
        this.findDocument();
        this.open();
        switch (this.action) {
        case START:
            this.start();
            break;
        case NEXT:
            this.gotoNextSlide();
            break;
        case PREVIOUS:
            this.gotoPreviousSlide();
            break;
        case STOP:
            this.stop();
            break;
        case CLOSE:
            this.stop();
            this.close();
            break;
        case GOTO:
            this.gotoSlide();
            break;
        }
    }
    out.print(this.showHTML());
%>
