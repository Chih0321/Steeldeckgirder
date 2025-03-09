import ezdxf

# create a new DXF R2010 document
doc = ezdxf.new("R2010")

# add new entities to the modelspace
msp = doc.modelspace()
# add a LINE entity
p1 = (0, 0)
msp.add_line(p1, (10, 0))
msp.add_line((10, 10), (10, 0))
# save the DXF document
doc.saveas("line.dxf")