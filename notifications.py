import application
import datetime
import sys
import win32com.client


#
# HELPER FUNCTIONS
#

# Returns the matching notebook from a list of Notebooks
# Returns None if the notebook is not found
def find_notebook_by_nickname(name, hierarchy):
    for n in hierarchy:
        if n.nickname == name:
            return n
    print ("The" + name +" notebook was not found")
    return None



# In Python 3.2 the strptime() handling of non-standard characters is broken - since we have a trailing Z, it breaks
#This method approximates by discarding microseconds & timezone - since we always report time in Z, this is a non-issue
#Refactor this when the bug in Python 3.2 is resolved

def parse_datetime(datetime_string):
    dt, foo, bar = datetime_string.partition(".")
    return datetime.datetime.strptime(dt, "%Y-%m-%dT%H:%M:%S")
    


# Checks to see if the given datetime stamp is newer than X seconds
# Returns the difference in seconds
# Default value is one day (86,400 seconds)

def is_newer_than(node, time=86400):
    diff = datetime.datetime.utcnow() - parse_datetime(node.last_modified_time)
    if (diff.total_seconds() < time):
        return diff
    else: 
        return None


# Recursively returns the first author it finds , walking up the hierarchy
def get_author_recursive(object):
    author = object.last_modified_by
    if ( author != "" and author != None):
        return author
    else:
        return(get_author_recursive(object.parent))



def get_changes_in_notebook(notebook_nickname):
    """Takes a notebook nickname & a reference to the hierarchy and returns changes within that notebook """
    nbk = find_notebook_by_nickname(notebook_nickname, on.hierarchy)
    return folder_handler(nbk)




#
# HTML FORMATTING HELPERS
#


def generate_margin(margin):
    """ Generates a margin from an integer value"""
    result = ""
    for x in range(0, margin):
        result += "&nbsp;&nbsp;&nbsp;&nbsp;"
    return result


def construct_folder_html(folder, changes):
    """ Takes in a folder object, integer value for margin &  a list of text changes & generates formatted HTML to place in an email"""
    #folder_link =  "<a href = '" + on.server.GetHyperlinkToObject(folder.id) + "' >" + folder.name + " </a> <br/>"
    #Enforcing Segoe UI font
    return ("<font face = 'Segoe UI'>" + changes + "</font>") 


def construct_breadcrumb(object):
    """Constructs HTML breadcrumb for sections"""
    if type(object.parent) == application.SectionGroup:
        link = "<a href = '" + on.server.GetHyperlinkToObject(object.parent.id) + "' >" + object.parent.name + "</a> / "
        return (link + construct_breadcrumb(object.parent))
    else:
        return ""


def construct_section_html(section, changes):
    section_link =  "<a href = '" + on.server.GetHyperlinkToObject(section.id) + "' >" + section.name + " </a><br/><br/>"
    return ("<font color = '#D9D9D9'>---------------------------------</font>" + "<br/>" + construct_breadcrumb(section) + section_link+ changes) 
        

def construct_page_html(page, changes):
    page_link = "<a href = '" + on.server.GetHyperlinkToObject(page.id) + "' >" + page.name + "</a> - "
    hours = round( is_newer_than(page).seconds / 3600 )
    page_changed_at =  "<span style='font-style: italic;'>%d %s ago </span>"  % (hours, "Hour" if hours==1 else "Hours")
    page_recent_changes =  "<font color = '#7F7F7F'>"+ changes + "</font><br/>"
    html_margin = generate_margin(1)
    
    return (" <span style='font-size: 12px;'>" + html_margin + page_link + "  " + page_changed_at + "<br/>" +
            html_margin + page_recent_changes + "</span><br/>")



#
# EMAIL HELPERS
#

def send_email(to, subject, body):
    """ Takes in recipients, subject and body and sends out an email using the global outlook process """

    FOOTER = ("<font color = '#D9D9D9' face = 'Segoe UI'>---------------------------------</font><br/>" +
              "<span style='font-size: 12px; font-family: Segoe UI; color: #7F7F7F;'>" +
              "This e-mail was automatically generated. To unsubscribe, get the " +
              "source, report bugs or request features, contact __person__ </span>")
    
    email = outlook_process.CreateItem(0)
    email.To = to
    email.Subject = subject
    email.HTMLBody = body + FOOTER
    email.Display()
        

def dispatch_emails (recipients, notebook_nickname):
    email_body = get_changes_in_notebook(notebook_nickname)
    if email_body != "":
        send_email(recipients, "Recent Changes in " + notebook_nickname, email_body)
    else:
        print("No changes were found")



#
# HANDLER FUNCTIONS
# These parse through different types of OneNote objects, letting us know if there were any changes
#

def folder_handler(folder):
    changes = ""
    
    for child in folder:
        if is_newer_than(child):
            if (type(child) == application.Section):
                changes += section_handler(child)
            elif (type(child) == application.SectionGroup):
                changes += folder_handler (child)



                
    if changes == "":
        return changes
    else:
        return (construct_folder_html(folder, changes))




# Section  - this returns a section string or a false if the pages contain no changes. Calls pages.
                
def section_handler(section):
    changes = ""
    
    for page in section:
        if is_newer_than(page):
            changes += page_handler(page)

            
    if changes == "":
        return changes
    else:
        return (construct_section_html(section, changes))



# Pages - this returns a page string or false if the pagecontent has no changes. Calls page content
def page_handler(page):
    changes = page_content_handler(page.id)

    if changes == "":
        return changes
    else:
        return (construct_page_html(page, changes)) 


# If the page is new, do we want to do something special here? We can check dateTime, and add it to a new list

def new_page_handler(page):
    pass
    



def page_content_handler(pageID):
    page = on.get_page_content(pageID)
    changes = 0
    authors = set()
    # Title could be a child of page, what do we wnt to do here?
    for child in page:
        if (type(child) == application.Outline):
            if is_newer_than(child):
                for oe in child:
                    c, a = count_oe_changes(oe)
                    changes += c
                    authors = a | authors

    if changes == 0:
        return ""
    else:
        result = "%d %s by "  % (changes,"Change" if changes==1 else "Changes")
        result+= ', '.join(authors)
        return result




def count_oe_changes(oe):
    """Returns the number of changes made within the OE, and a set of unique authors for those changes """
    changes = 0
    authors = set()

    if is_newer_than(oe):
        changes += 1
        authors.add(get_author_recursive(oe))

    for child in oe:
        r,a = count_oe_changes(child)
        changes += r
        authors = a | authors

    return changes, authors




#Main

on = application.OneNote()
outlook_process = win32com.client.gencache.EnsureDispatch("Outlook.Application.14")

