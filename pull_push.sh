NOW=$(date +"%Y-%m-%d-%H-%M-%S")
commitMsg="Checkpoint commit as of ${NOW}."

git add -A && git commit -am"${commitMsg}"
git pull
git push

echo "waiting 15 seconds"
sleep 15s