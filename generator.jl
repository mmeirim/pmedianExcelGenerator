using FacilityLocationProblems
using XLSX

file_path = "excel_files/"
instances = getPMedianInstances()

function generateSheet(sheet, data, name)
    XLSX.rename!(sheet, name)
    sheet["E1"] = "Number of medians (p)"
    sheet["F1"] = data.medians
    sheet["E2"] = "Medians capacities"
    sheet["F2"] = data.capacity

    sheet["A1"] = "X"
    sheet["B1"] = "Y"
    sheet["C1"] = "Demands"

    sheet["A2", dim=1] = data.x
    sheet["B2", dim=1] = data.y
    sheet["C2", dim=1]= data.demands
end

groups = [ "A", "B", "C", "E", "F", "H", "K", "L", "N", "O" ]
for i in 1:10
    file_name = joinpath(file_path, "group" * groups[i] * ".xlsx")
    XLSX.openxlsx(file_name, mode = "w") do xf
        XLSX.addsheet!(xf)

        data1 = loadPMedianProblem(instances[i])
        sheet1 = xf[1]
        generateSheet(sheet1, data1, "Case 1")

        data2 = loadPMedianProblem(instances[i + 10])
        sheet2 = xf[2]
        generateSheet(sheet2, data2, "Case 2")
    end
end