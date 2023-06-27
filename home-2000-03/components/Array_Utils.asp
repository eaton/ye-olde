<%

'--------------------------------------------------------------------------------
' The ArrayManip Object brings similar array manipulation methods to VBScript as 
' found in other languages such as JScript and Perl. Most developers will be 
' immediately familiar with these new methods as they work just like their Perl 
' and JScript equivalents. There are currently ten methods exposed by this class. 
' Most of these methods directly modify the input array and return information 
' about the array, not the array itself. In the event that the input array is 
' returned by a method's execution, the original array will be modified as well 
' so the output of those methods can be discarded if necessary.
' 
' Docs at http://www.aspEmporium.com/aspEmporium/examples/arraymanip_class.asp
' 
' Code ©2000 ASP Emporium, http://www.aspEmporium.com
'--------------------------------------------------------------------------------


Class ArrayManip

	 ' useful array manipulation functions 
	 ' for vbscript arrays

	Public Function Pop(byRef theArray)
		 ' returns the last value in the 
		 ' array and removes it from the 
		 ' array, shortening the array
		 ' by one element

		Pop = theArray(UBound(theArray))
		Redim Preserve theArray(UBound(theArray) - 1)
	End Function

	Public Function Push(byRef theArray, byVal theDataToAppend)
		 ' appends new elements to an 
		 ' array and returns the new 
		 ' length of the array (ubound)
		dim itemCount, tmp, i, oldUBound, j
		oldUBound = UBound(theArray)
		itemCount = 0
		tmp = split(theDataToAppend, ",")
		itemCount = UBound(tmp)
		Redim Preserve theArray(oldUBound + itemCount)
		i = 0
		for j = oldUBound + 1 to ubound(theArray)
			theArray(j) = trim(tmp(i))
			i = i + 1
		next
		Push = UBound(theArray)
	End Function

	Public Function Shift(byRef theArray)
		 ' removes the first element of an array
		 ' and displays it. Shifts every other element
		 ' down one element and shortens the array by 
		 ' 1 element.

		dim i
		Shift = theArray(LBound(theArray))
		for i = 1 to ubound(theArray)
			theArray(i - 1) = theArray(i)
		next
		Redim Preserve theArray(UBound(theArray) - 1)
	End Function

	Public Function Splice(byRef theArray, byVal start, byVal deletecount, byVal optionalList)
		 ' removes elements from an array and 
		 ' optionally inserts new values to
		 ' replace the deleted elements.
		 ' Returns the removed elements as a 
		 ' new array.

		dim i, j, newArray()
		dim tmp, outputArray()
		dim ct, arrayUB
		arrayUB = ubound(theArray)
		ct = 0 : j = 0
		if (deletecount < 0) or _
			(deletecount > arrayUB) or _
			(not isnumeric(deletecount)) then _
		deletecount = arrayUB
		if (start < 0) or (not isnumeric(start)) then start = 0
		if start > arrayUB then _
			start = arrayUB

		redim newArray(deletecount)
		for i = start to (start + deletecount - 1)
			newArray(j) = theArray(i)
			theArray(i) = ""
			j = j + 1
		next
		tmp = split(optionalList, ",")
		j = start
		for i = 0 to ubound(tmp)
			theArray(j) = trim(tmp(i))
			if j = (start + deletecount - 1) then 
				exit for
			end if
			j = j + 1
		next
		Splice = newArray
		for i = 0 to arrayUB
			theArray(i) = trim(theArray(i))
			if Len(theArray(i)) = 0 then
				ct = ct + 1
			end if
		next
		Redim outputArray(  arrayUB - ct  )
		j = 0
		for i = 0 to ubound(theArray)
			if Not Len(trim(theArray(i))) = 0 then
				outputArray(j) = theArray(i)
				j = j + 1
			end if
		next
		theArray = outputArray
	End Function

	Public Function UnShift(byRef theArray, byVal theDataToPrepend)
		 ' returns an array with the specified 
		 ' elements added to the beginning of
		 ' the original array
		dim tmp, i, newArray()
		dim j
		tmp = split(theDataToPrepend, ",")
		Redim newArray(ubound(theArray) + ubound(tmp) + 1)
		j = ubound(tmp) + 1
		for i = 0 to ubound(theArray)
			newArray(j + i) = theArray(i)
		next
		for i = 0 to ubound(tmp)
			newArray(i) = trim(tmp(i))
		next
		UnShift = newArray
		theArray = newArray
	End Function

	Public Function HasDups(byRef theArray)
		dim d, item, bER
		bER = false
		set d = createobject("scripting.dictionary")
		on error resume next
		for each item in theArray
			d.add item, ""
			if Err then 
				bER = true
				exit for
			end if
		next
		on error goto 0
		d.removeall
		set d = nothing
		HasDups = bER
	End Function

	Public Function RemDups(byRef theArray)
		dim d, item, bER, newArray()
		dim i, a
		i = 0
		redim newArray(ubound(theArray))
		bER = false
		set d = createobject("scripting.dictionary")
		on error resume next
		for each item in theArray
			d.add item, ""
		next
		on error goto 0
		a = d.keys
		d.removeall
		set d = nothing
		RemDups = a
		theArray = a
	End Function

	Public Function RevArray(byRef arrayinput)
		Dim i, ubnd
		Dim newarray()
		ubnd = UBound( arrayinput )
		Redim newarray(ubnd)
		For i = 0 to UBound( arrayinput )
			newarray( ubnd - i ) = arrayinput( i )
		Next
		RevArray = newarray
		arrayInput = newarray
	End Function

	Public Function Sort(byRef unsortedarray)
		Dim Front, Back, Loc, Temp, Arrsize
		Arrsize = UBOUND(unsortedarray)
			For Front = 0 To Arrsize - 1
				Loc = Front
				For Back = Front To Arrsize
					If unsortedarray(Loc) > _
					    unsortedarray(Back) Then
						Loc = Back
					End If
				Next
				Temp = unsortedarray(Loc)
				unsortedarray(Loc) = unsortedarray(Front)
				unsortedarray(Front) = Temp
			Next
		Sort = unsortedarray
	End Function

	Public Function Slice(byRef theArray, byVal start, byVal theend)
		 ' returns part of an array as a new array. 
		 ' doesn't modify the original array
		dim lstart, lend, i, j, newArray()
		lstart = lbound(theArray) : lend = ubound(theArray)
		if start < lstart then start = 0
		if start > lend then start = lend
		if theend > lend then theend = lend
		if theend < lstart then theend = 0
		if theend = "" then theend = lend
		redim preserve newArray(theend - start)
		j = 0
		for i = start to theend
			newArray(j) = theArray(i) : j = j + 1
		next
		Slice = newArray
	End Function
End Class
%>
