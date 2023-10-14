using FacilityLocationProblems
using XLSX


file_path = "excel_files/"
instances = getPMedianInstances()

function generateSheet(sheet, data)
    XLSX.rename!(sheet, data.name)
    sheet["A1"] = "Number of medians (p)"
    sheet["B1"] = data.medians
    sheet["A2"] = "Medians capacities"
    sheet["B2"] = data.capacity

    sheet["A4"] = "X"
    sheet["B4"] = "Y"
    sheet["C4"] = "Demands"
    sheet["E4"] = "X | Y"
    sheet["E3"] = "Costs matrix (distances)"
    sheet["E5", dim=1] = collect(1:length(data.x))
    sheet["F4"] = collect(1:length(data.y))

    sheet["A5", dim=1] = data.x
    sheet["B5", dim=1] = data.y
    sheet["C5", dim=1]= data.demands
    sheet["F5"]= data.costs
end

for idx in 1:10
    XLSX.openxlsx(file_path * "group" * string(idx) * ".xlsx", mode="w") do xf
        XLSX.addsheet!(xf)

        data1 = loadPMedianProblem(instances[idx])
        data2 = loadPMedianProblem(instances[idx+10])

        sheet1 = xf[1]
        sheet2 = xf[2]

        generateSheet(sheet1, data1)
        generateSheet(sheet2, data2)
    end
end