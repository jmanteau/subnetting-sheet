import xlsxwriter
import netaddr
import string


# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('subnet.xlsx')
worksheet = workbook.add_worksheet()

header_format = workbook.add_format({'align': 'center',
                                     'valign': 'vcenter',
                                     'font_color': 'white',
                                     'bg_color': 'gray',
                                     'bold': True,
                                     'border': 1,
                                     'text_wrap': True})

cell_format = workbook.add_format({'align': 'center',
                                   'valign': 'vcenter',
                                   'border': 1,
                                   'text_wrap': True})

# The netmask to which we want to go
smallestmask = 23

# The original subnet to split
net = netaddr.IPNetwork('10.84.0.0/16')

# How many columns are required
nbcolumns = smallestmask - net.prefixlen

# The number of subnets the smallest mask requires
maxlines = len(list(net.subnet(smallestmask)))

# Change column size
worksheet.set_column(0,nbcolumns,20)

# For the number of netmask
for col in range(0,nbcolumns+1):

    # get the letter for the column
    colletter=string.ascii_uppercase[col]

    # header writing
    firstnet= list(net.subnet(net.prefixlen+col))[0]
    headercell= "%s1" % colletter
    header= "/%s - %s hosts \n %s nets %s" % (str(net.prefixlen+col), str(len(firstnet)),
                                              str(len(list(net.subnet(net.prefixlen+col)))), str(firstnet.netmask) )

    worksheet.write(headercell, header, header_format)

    # get the size of one merged cell for this netmask
    sizecell = maxlines/2**col

    # for each subnet corresponding to the split for this netmask of the larger subnet
    for i, snet in enumerate(net.subnet(net.prefixlen+col)):
        # Beginning Cell
        bcell= "%s%s" % (colletter,int(2+i*sizecell))
        # End Cell
        ecell= "%s%s" % (colletter,int(2+i*sizecell+sizecell-1))


        if col == nbcolumns:
            # last column, do not merge cells
            worksheet.write(bcell, str(snet), cell_format)
        else:
            # We can only write simple types to merged ranges so we write a blank string.
            # We merge from bcell to ecell
            worksheet.merge_range("%s:%s" % (bcell,ecell), "" , cell_format)

            # We then overwrite the first merged cell with a rich string. Note that we
            # must also pass the cell format used in the merged cells format at the end.
            worksheet.write(bcell, str(snet), cell_format)


workbook.close()
