async function getGraphToken() {
    return new Promise((resolve, reject) => {
        Office.auth.getAccessTokenAsync({ forceConsent: false }, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}
