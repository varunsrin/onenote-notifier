import notifications as notif
import sys

#Arg 1 is the list of emails, Arg 2 is the name of the notebook

USER_NAME = "User Foo"

def main():
    EMAIL_LIST = sys.argv[1]
    NBK_NAME = sys.argv[2]
    notif.dispatch_emails (EMAIL_LIST, NBK_NAME, USER_NAME)

if __name__ == "__main__":
    main()


