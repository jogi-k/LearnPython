# Write your code here :-)

# Aufgabe in Mathe für Laurin
# Ziehe zwei Zahlen voneinander ab.
# Die beiden Zahlen müssen dabei jede der Ziffern 1-8 genau einmal enthalten.
# Die erhaltene Differenz soll so nahe wie möglich bei 5000 sein.


def check_unique( number ):
    mylist = list(str(number))
    if '0' in mylist :
        return False
    if '9' in mylist :
        return False
    myset  = set( mylist )
    mylist2 = sorted( list( myset ) )
    mylist3 = sorted( mylist )
    if mylist2 == mylist3 :
        return True
    return False

minimum = 9000
for x in range ( 10000000 , 90000000 ):
    if check_unique( x ) :
        num1 = x // 10000
        num2 = x % 10000
        res = num1 - num2;
        diff = res - 5000
        absdiff = abs( diff )

        print ( format(absdiff,"0>5"), format(diff," >6") , format(res, " >5") ,  num1, num2 )
        if absdiff < minimum :
            minimum = absdiff
            minres  = res
            minnum1 = num1
            minnum2 = num2
print ( "Minimum gefunden : " , minnum1 , "-" , minnum2 , "=" , minres, "Abstand zu 5000:" , minimum )
