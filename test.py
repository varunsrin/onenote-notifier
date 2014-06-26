# Testing_ON.py
import onepy

on = onepy.OneNote()

for n in on.hierarchy:
    print (n)


    # for s in n:
    #     print("   " + s.name)
    #     for p in s:
    #         print ("                   " + p.name)
    #         print ("                   " + p.id)
    #         print ("                   " + on.process.GetHyperlinkToObject(p.id))



# PAGE CONTENT TESTING

#f = on.get_page_content("{870EB865-299E-4A78-8B3B-ED88EE3CF0C4}{1}{B0}")
#
#for o in f:
#    print(o)
#    for oe in o:
#        try:
#            print ("   " + oe.text)
#        except:
#            print("failure to print oe")
#
#
#
