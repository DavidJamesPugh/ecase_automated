pushd \\SAV-FP01\data$\Shared Data\Care Manager\Quality Data\Data Technician\eCase Migration
pip freeze > old_requirements.txt
pip freeze > requirements.txt
pur -r requirements.txt

popd